import { NextRequest, NextResponse } from 'next/server';
import { requireSession } from '@/lib/auth/session';
import { prisma } from '@/lib/db/client';
import { decryptToken, encryptToken } from '@/lib/crypto/token';
import {
  fetchBoardItems,
  fetchBoardColumns,
  fetchMondayUsers,
  findPeopleColumn,
  createMondayItem,
  updateMondayItem,
  type MondayItem,
  type MondayUser,
} from '@/lib/monday/api';
import { createEvent, updateEvent, listEvents, type CalendarEvent } from '@/lib/ms/graph';
import { refreshAccessToken } from '@/lib/ms/oauth';

function toIsoDate(d: Date) {
  return d.toISOString().slice(0, 10);
}

function toIsoTime(d: Date) {
  return d.toISOString().slice(11, 19);
}

function msEventToStartEnd(event: CalendarEvent): { start: Date; end: Date } {
  return {
    start: new Date(event.start.dateTime),
    end: new Date(event.end.dateTime),
  };
}

function buildMondayColumnValues(
  item: MondayItem,
  columns: Awaited<ReturnType<typeof fetchBoardColumns>>,
  opts: {
    start: Date;
    end: Date;
    subject: string;
    location?: string;
    attendees?: Array<{ id: number; kind: 'person' }>;
    peopleColId?: string;
  },
): Record<string, unknown> {
  const dateCol = columns.find((c) => c.type === 'date');
  const durationCol = columns.find((c) => c.type === 'numbers');
  const textCol = columns.find((c) => c.type === 'text');

  const values: Record<string, unknown> = {};

  if (dateCol) {
    const durationH =
      durationCol
        ? Math.max(0, (opts.end.getTime() - opts.start.getTime()) / 3600_000)
        : 0;
    const endDate = durationCol
      ? new Date(opts.start.getTime() + durationH * 3600_000)
      : opts.end;
    values[dateCol.id] = {
      date: toIsoDate(opts.start),
      time: toIsoTime(opts.start),
    };
    // If board also has a timeline or second date column for end date we would need it;
    // for a single date column with duration we just send date+time and duration separately.
  }

  if (durationCol && dateCol) {
    const durationH = Math.max(0, (opts.end.getTime() - opts.start.getTime()) / 3600_000);
    values[durationCol.id] = durationH;
  }

  if (textCol && opts.location) {
    values[textCol.id] = opts.location;
  }

  if (opts.peopleColId && opts.attendees) {
    values[opts.peopleColId] = { personsAndTeams: opts.attendees };
  }

  return values;
}

export async function POST(req: NextRequest) {
  const session = await requireSession(req);
  const boardId = process.env.MONDAY_BOARD_ID;
  if (!boardId) {
    return NextResponse.json({ error: 'MONDAY_BOARD_ID not configured' }, { status: 500 });
  }

  const mondayBoardId = BigInt(boardId);

  const [mondayAcc, msAcc] = await Promise.all([
    prisma.mondayAccount.findUnique({ where: { userId: session.userId } }),
    prisma.msAccount.findUnique({ where: { userId: session.userId } }),
  ]);

  if (!mondayAcc) {
    return NextResponse.json({ error: 'monday not connected' }, { status: 400 });
  }
  if (!msAcc) {
    return NextResponse.json({ error: 'microsoft not connected' }, { status: 400 });
  }

  const mondayToken = decryptToken(mondayAcc.accessTokenEnc);
  let msToken = decryptToken(msAcc.accessTokenEnc);

  // Refresh MS token if expired
  if (msAcc.expiresAt < new Date()) {
    const refresh = decryptToken(msAcc.refreshTokenEnc);
    const tok = await refreshAccessToken(refresh);
    msToken = tok.access_token;
    await prisma.msAccount.update({
      where: { userId: session.userId },
      data: {
        accessTokenEnc: encryptToken(tok.access_token),
        refreshTokenEnc: encryptToken(tok.refresh_token),
        expiresAt: new Date(Date.now() + tok.expires_in * 1000),
      },
    });
  }

  // Time window: 7 days ago → 90 days from now
  const windowStart = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  const windowEnd = new Date(Date.now() + 90 * 24 * 60 * 60 * 1000);

  // Fetch board columns, monday items, and outlook events
  const [columns, mondayItems, outlookEventsResp] = await Promise.all([
    fetchBoardColumns(mondayToken, mondayBoardId),
    fetchBoardItems(mondayToken, mondayBoardId),
    listEvents(msToken, {
      calendarId: msAcc.selectedCalendarId ?? undefined,
      start: windowStart,
      end: windowEnd,
    }),
  ]);

  const outlookEvents = outlookEventsResp.value;
  const itemsWithDate = mondayItems.filter(
    (i) => i.dateStart && i.dateEnd && i.dateStart >= windowStart && i.dateStart <= windowEnd,
  );

  const results = {
    mondayToOutlook: { created: 0, updated: 0, failed: 0 },
    outlookToMonday: { created: 0, updated: 0, failed: 0 },
    conflicts: 0,
    attendeesSkipped: 0,
  };

  // ─── Attendee resolution setup ───
  const peopleCol = findPeopleColumn(columns);
  let userByEmail: Map<string, MondayUser> | null = null;
  let userById: Map<string, MondayUser> | null = null;
  if (peopleCol) {
    try {
      const users = await fetchMondayUsers(mondayToken);
      userByEmail = new Map();
      userById = new Map();
      for (const u of users) {
        if (u.email) userByEmail.set(u.email.toLowerCase(), u);
        userById.set(u.id.toString(), u);
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      await prisma.syncLog.create({
        data: {
          userId: session.userId,
          direction: 'bidirectional',
          action: 'fetch-users-failed',
          message: `monday users query failed; attendee sync disabled: ${msg}`,
        },
      });
      userByEmail = null;
      userById = null;
    }
  }

  function resolveMondayAttendees(item: MondayItem): string[] {
    if (!userById) return [];
    const emails: string[] = [];
    for (const person of item.attendees) {
      const user = userById.get(person.id.toString());
      if (user && user.email) {
        emails.push(user.email.toLowerCase());
      } else {
        results.attendeesSkipped++;
      }
    }
    return emails;
  }

  function resolveOutlookAttendees(
    event: CalendarEvent,
  ): Array<{ id: number; kind: 'person' }> {
    if (!userByEmail) return [];
    const persons: Array<{ id: number; kind: 'person' }> = [];
    for (const att of event.attendees ?? []) {
      const email = att.emailAddress?.address?.toLowerCase();
      if (!email) {
        results.attendeesSkipped++;
        continue;
      }
      const user = userByEmail.get(email);
      if (user) {
        persons.push({ id: Number(user.id), kind: 'person' });
      } else {
        results.attendeesSkipped++;
      }
    }
    return persons;
  }

  // Build lookup maps
  const mondayItemMap = new Map<string, MondayItem>();
  for (const item of mondayItems) mondayItemMap.set(item.id, item);

  const outlookEventMap = new Map<string, CalendarEvent>();
  for (const ev of outlookEvents) outlookEventMap.set(ev.id, ev);

  // Fetch existing mappings
  const mappings = await prisma.eventMapping.findMany({
    where: { userId: session.userId, deletedAt: null },
  });

  const mappingByMondayId = new Map<string, typeof mappings[0]>();
  const mappingByGraphId = new Map<string, typeof mappings[0]>();
  for (const m of mappings) {
    mappingByMondayId.set(m.mondayItemId.toString(), m);
    mappingByGraphId.set(m.graphEventId, m);
  }

  // Helper to update mapping etags
  async function updateMappingEtags(
    mappingId: string,
    mondayEtag: string,
    graphEtag: string,
  ) {
    await prisma.eventMapping.update({
      where: { id: mappingId },
      data: { mondayEtag, graphEtag, lastSyncedAt: new Date() },
    });
  }

  // ─── 1) Process existing mappings (bidirectional) ───
  for (const mapping of mappings) {
    const mondayItem = mondayItemMap.get(mapping.mondayItemId.toString());
    const outlookEvent = outlookEventMap.get(mapping.graphEventId);

    if (!mondayItem || !outlookEvent) {
      // One side was deleted; skip for now (deletion sync can be added later)
      continue;
    }

    const mondayChanged = mondayItem.updatedAt !== mapping.mondayEtag;
    const outlookChanged = (outlookEvent.lastModifiedDateTime ?? '') !== mapping.graphEtag;

    if (!mondayChanged && !outlookChanged) continue;

    try {
      if (mondayChanged && outlookChanged) {
        // Conflict: pick the newer one
        const mondayTime = new Date(mondayItem.updatedAt).getTime();
        const outlookTime = new Date(outlookEvent.lastModifiedDateTime ?? 0).getTime();
        if (mondayTime >= outlookTime) {
          // monday wins
          const body = `Synced from monday.com item <a href="https://monday.com/boards/${mondayItem.boardId}/pulses/${mondayItem.id}">${mondayItem.name}</a>`;
          await updateEvent(msToken, outlookEvent.id, {
            subject: mondayItem.name,
            start: mondayItem.dateStart!,
            end: mondayItem.dateEnd!,
            body,
            location: mondayItem.location ?? undefined,
            attendees: userById ? resolveMondayAttendees(mondayItem) : undefined,
          });
          const updatedEvent = await listEvents(msToken, {
            calendarId: msAcc.selectedCalendarId ?? undefined,
          }).then((r) => r.value.find((e) => e.id === outlookEvent.id));
          await updateMappingEtags(
            mapping.id,
            mondayItem.updatedAt,
            updatedEvent?.lastModifiedDateTime ?? outlookEvent.lastModifiedDateTime ?? '',
          );
          results.mondayToOutlook.updated++;
        } else {
          // outlook wins
          const { start, end } = msEventToStartEnd(outlookEvent);
          const colValues = buildMondayColumnValues(mondayItem, columns, {
            start,
            end,
            subject: outlookEvent.subject,
            location: outlookEvent.location?.displayName,
            attendees: userByEmail ? resolveOutlookAttendees(outlookEvent) : undefined,
            peopleColId: userByEmail ? peopleCol?.id : undefined,
          });
          const updated = await updateMondayItem(
            mondayToken,
            mondayBoardId,
            mondayItem.id,
            colValues,
          );
          await updateMappingEtags(mapping.id, updated.updated_at, outlookEvent.lastModifiedDateTime ?? '');
          results.outlookToMonday.updated++;
        }
        results.conflicts++;
      } else if (mondayChanged) {
        const body = `Synced from monday.com item <a href="https://monday.com/boards/${mondayItem.boardId}/pulses/${mondayItem.id}">${mondayItem.name}</a>`;
        await updateEvent(msToken, outlookEvent.id, {
          subject: mondayItem.name,
          start: mondayItem.dateStart!,
          end: mondayItem.dateEnd!,
          body,
          location: mondayItem.location ?? undefined,
          attendees: userById ? resolveMondayAttendees(mondayItem) : undefined,
        });
        const updatedEvent = await listEvents(msToken, {
          calendarId: msAcc.selectedCalendarId ?? undefined,
        }).then((r) => r.value.find((e) => e.id === outlookEvent.id));
        await updateMappingEtags(
          mapping.id,
          mondayItem.updatedAt,
          updatedEvent?.lastModifiedDateTime ?? outlookEvent.lastModifiedDateTime ?? '',
        );
        results.mondayToOutlook.updated++;
      } else if (outlookChanged) {
        const { start, end } = msEventToStartEnd(outlookEvent);
        const colValues = buildMondayColumnValues(mondayItem, columns, {
          start,
          end,
          subject: outlookEvent.subject,
          location: outlookEvent.location?.displayName,
          attendees: userByEmail ? resolveOutlookAttendees(outlookEvent) : undefined,
          peopleColId: userByEmail ? peopleCol?.id : undefined,
        });
        const updated = await updateMondayItem(
          mondayToken,
          mondayBoardId,
          mondayItem.id,
          colValues,
        );
        await updateMappingEtags(mapping.id, updated.updated_at, outlookEvent.lastModifiedDateTime ?? '');
        results.outlookToMonday.updated++;
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      await prisma.syncLog.create({
        data: {
          userId: session.userId,
          direction: 'bidirectional',
          action: 'sync-mapping-failed',
          mappingId: mapping.id,
          message: `monday ${mapping.mondayItemId} / outlook ${mapping.graphEventId}: ${msg}`,
        },
      });
      results.mondayToOutlook.failed++;
    }
  }

  // ─── 2) New monday items → create Outlook events ───
  for (const item of itemsWithDate) {
    if (mappingByMondayId.has(item.id)) continue;

    try {
      const body = `Synced from monday.com item <a href="https://monday.com/boards/${item.boardId}/pulses/${item.id}">${item.name}</a>`;
      const calendarId = msAcc.selectedCalendarId ?? undefined;
      const event = await createEvent(msToken, {
        subject: item.name,
        start: item.dateStart!,
        end: item.dateEnd!,
        body,
        location: item.location ?? undefined,
        calendarId,
        attendees: userById ? resolveMondayAttendees(item) : undefined,
      });
      await prisma.eventMapping.create({
        data: {
          userId: session.userId,
          mondayItemId: BigInt(item.id),
          mondayBoardId: item.boardId,
          graphEventId: event.id,
          graphCalendarId: calendarId || 'default',
          mondayEtag: item.updatedAt,
          graphEtag: event.lastModifiedDateTime ?? '',
          origin: 'monday',
          lastSyncedAt: new Date(),
        },
      });
      results.mondayToOutlook.created++;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      await prisma.syncLog.create({
        data: {
          userId: session.userId,
          direction: 'monday→outlook',
          action: 'create-event-failed',
          message: `item ${item.id}: ${msg}`,
        },
      });
      results.mondayToOutlook.failed++;
    }
  }

  // ─── 3) New Outlook events → create monday items ───
  for (const event of outlookEvents) {
    if (mappingByGraphId.has(event.id)) continue;

    try {
      const { start, end } = msEventToStartEnd(event);
      const colValues = buildMondayColumnValues(
        {} as MondayItem,
        columns,
        {
          start,
          end,
          subject: event.subject,
          location: event.location?.displayName,
          attendees: userByEmail ? resolveOutlookAttendees(event) : undefined,
          peopleColId: userByEmail ? peopleCol?.id : undefined,
        },
      );
      const item = await createMondayItem(
        mondayToken,
        mondayBoardId,
        event.subject || '(no subject)',
        colValues,
      );
      await prisma.eventMapping.create({
        data: {
          userId: session.userId,
          mondayItemId: BigInt(item.id),
          mondayBoardId,
          graphEventId: event.id,
          graphCalendarId: msAcc.selectedCalendarId || 'default',
          mondayEtag: item.updated_at,
          graphEtag: event.lastModifiedDateTime ?? '',
          origin: 'outlook',
          lastSyncedAt: new Date(),
        },
      });
      results.outlookToMonday.created++;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      await prisma.syncLog.create({
        data: {
          userId: session.userId,
          direction: 'outlook→monday',
          action: 'create-item-failed',
          message: `event ${event.id}: ${msg}`,
        },
      });
      results.outlookToMonday.failed++;
    }
  }

  return NextResponse.json({
    totalMondayItems: mondayItems.length,
    totalOutlookEvents: outlookEvents.length,
    mappings: mappings.length,
    ...results,
  });
}
