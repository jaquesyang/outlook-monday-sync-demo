const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

async function graphFetch<T>(accessToken: string, path: string, opts?: RequestInit): Promise<T> {
  const r = await fetch(`${GRAPH_BASE}${path}`, {
    ...opts,
    headers: {
      authorization: `Bearer ${accessToken}`,
      'content-type': 'application/json',
      ...(opts?.headers || {}),
    },
  });
  if (!r.ok) {
    const text = await r.text();
    throw new Error(`graph request failed: ${r.status} ${text}`);
  }
  return r.json() as Promise<T>;
}

export type CalendarEvent = {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  body?: { contentType: string; content: string };
  location?: { displayName: string };
  lastModifiedDateTime?: string;
  attendees?: Array<{
    emailAddress: { address: string; name?: string };
    type?: 'required' | 'optional' | 'resource';
    status?: { response: string; time?: string };
  }>;
};

function toGraphAttendees(emails: string[] | undefined) {
  if (!emails) return undefined;
  return emails.map((email) => ({
    emailAddress: { address: email },
    type: 'required' as const,
  }));
}

export async function createEvent(
  accessToken: string,
  opts: {
    subject: string;
    start: Date;
    end: Date;
    body?: string;
    location?: string;
    calendarId?: string;
    attendees?: string[];
  },
): Promise<CalendarEvent> {
  const path = opts.calendarId
    ? `/me/calendars/${opts.calendarId}/events`
    : '/me/events';

  // Graph API requires end > start for regular events.
  let end = opts.end;
  if (end.getTime() <= opts.start.getTime()) {
    end = new Date(opts.start.getTime() + 60 * 60 * 1000); // default 1 hour
  }

  return graphFetch<CalendarEvent>(accessToken, path, {
    method: 'POST',
    body: JSON.stringify({
      subject: opts.subject,
      start: { dateTime: opts.start.toISOString(), timeZone: 'UTC' },
      end: { dateTime: end.toISOString(), timeZone: 'UTC' },
      body: opts.body
        ? { contentType: 'html', content: opts.body }
        : undefined,
      location: opts.location
        ? { displayName: opts.location }
        : undefined,
      attendees: toGraphAttendees(opts.attendees),
    }),
  });
}

export async function updateEvent(
  accessToken: string,
  eventId: string,
  opts: {
    subject?: string;
    start?: Date;
    end?: Date;
    body?: string;
    location?: string;
    attendees?: string[];
  },
): Promise<CalendarEvent> {
  // Graph API requires end > start for regular events.
  let end = opts.end;
  if (opts.start !== undefined && end !== undefined && end.getTime() <= opts.start.getTime()) {
    end = new Date(opts.start.getTime() + 60 * 60 * 1000);
  }

  return graphFetch<CalendarEvent>(accessToken, `/me/events/${eventId}`, {
    method: 'PATCH',
    body: JSON.stringify({
      ...(opts.subject !== undefined && { subject: opts.subject }),
      ...(opts.start !== undefined && {
        start: { dateTime: opts.start.toISOString(), timeZone: 'UTC' },
      }),
      ...(end !== undefined && {
        end: { dateTime: end.toISOString(), timeZone: 'UTC' },
      }),
      ...(opts.body !== undefined && {
        body: { contentType: 'html', content: opts.body },
      }),
      ...(opts.location !== undefined && {
        location: opts.location ? { displayName: opts.location } : null,
      }),
      ...(opts.attendees !== undefined && {
        attendees: toGraphAttendees(opts.attendees) ?? [],
      }),
    }),
  });
}

export async function listEvents(
  accessToken: string,
  opts?: { calendarId?: string; start?: Date; end?: Date },
): Promise<{ value: CalendarEvent[] }> {
  const path = opts?.calendarId
    ? `/me/calendars/${opts.calendarId}/events`
    : '/me/events';
  const params = new URLSearchParams({
    $select: 'id,subject,start,end,body,location,lastModifiedDateTime,attendees',
    $top: '100',
  });
  if (opts?.start && opts?.end) {
    const filter = `start/dateTime ge '${opts.start.toISOString()}' and end/dateTime le '${opts.end.toISOString()}'`;
    params.set('$filter', filter);
  }
  return graphFetch<{ value: CalendarEvent[] }>(accessToken, `${path}?${params.toString()}`);
}
