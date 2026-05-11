const MONDAY_API = 'https://api.monday.com/v2';

export type MondayItem = {
  id: string;
  name: string;
  boardId: bigint;
  updatedAt: string;
  dateStart: Date | null;
  dateEnd: Date | null;
  location: string | null;
  durationHours: number | null;
  columnValues: Record<string, unknown>;
  attendees: Array<{ id: bigint; kind: string }>;
};

export type MondayUser = {
  id: bigint;
  name: string;
  email: string;
};

async function queryMonday<T>(accessToken: string, query: string, variables?: Record<string, unknown>): Promise<T> {
  const r = await fetch(MONDAY_API, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      authorization: accessToken,
      'API-Version': '2024-01',
    },
    body: JSON.stringify({ query, variables }),
  });
  if (!r.ok) throw new Error(`monday query failed: ${r.status} ${await r.text()}`);
  const j = await r.json() as { data?: T; errors?: Array<{ message: string }> };
  if (j.errors?.length) throw new Error(`monday query error: ${j.errors[0].message}`);
  return j.data!;
}

export async function fetchBoardItems(accessToken: string, boardId: bigint): Promise<MondayItem[]> {
  const query = `
    query ($boardIds: [ID!]) {
      boards(ids: $boardIds) {
        items_page(limit: 500) {
          items {
            id
            name
            updated_at
            column_values {
              id
              type
              value
            }
          }
        }
      }
    }
  `;
  type Resp = {
    boards: Array<{
      items_page: {
        items: Array<{
          id: string;
          name: string;
          updated_at: string;
          column_values: Array<{
            id: string;
            type: string;
            value: string | null;
          }>;
        }>;
      };
    }>;
  };
  const data = await queryMonday<Resp>(accessToken, query, { boardIds: [boardId.toString()] });
  const board = data.boards[0];
  if (!board) return [];

  return board.items_page.items.map((item) => {
    let dateStart: Date | null = null;
    let dateEnd: Date | null = null;
    let location: string | null = null;
    let durationHours: number | null = null;
    let attendees: Array<{ id: bigint; kind: string }> = [];
    const columnValues: Record<string, unknown> = {};

    for (const cv of item.column_values) {
      if (!cv.value) continue;
      try {
        const parsed = JSON.parse(cv.value);
        columnValues[cv.id] = parsed;

        if (cv.type === 'date' && parsed.date) {
          const time = parsed.time || '00:00:00';
          dateStart = new Date(`${parsed.date}T${time}`);
        }
        if (cv.type === 'timeline' && parsed.from && parsed.to) {
          dateStart = new Date(`${parsed.from}T00:00:00`);
          dateEnd = new Date(`${parsed.to}T23:59:59`);
        }
        if (cv.type === 'text' && typeof parsed === 'string') {
          location = parsed || null;
        }
        if (cv.type === 'numbers' && parsed !== null) {
          durationHours = Number(parsed);
        }
        if (
          (cv.type === 'people' || cv.id.startsWith('multiple_person_')) &&
          parsed &&
          Array.isArray(parsed.personsAndTeams)
        ) {
          attendees = parsed.personsAndTeams.map(
            (p: { id: number | string; kind?: string }) => ({
              id: BigInt(p.id),
              kind: p.kind ?? 'person',
            }),
          );
        }
      } catch {
        columnValues[cv.id] = cv.value;
        if (cv.type === 'text') {
          location = cv.value;
        }
      }
    }

    // Apply duration to compute end time if we have a date but no explicit end
    if (dateStart && !dateEnd && durationHours && durationHours > 0) {
      dateEnd = new Date(dateStart.getTime() + durationHours * 60 * 60 * 1000);
    }
    // Fallback: default to 1-hour event so Graph API accepts it (end must be > start)
    if (dateStart && !dateEnd) {
      dateEnd = new Date(dateStart.getTime() + 60 * 60 * 1000);
    }

    return {
      id: item.id,
      name: item.name,
      boardId,
      updatedAt: item.updated_at,
      dateStart,
      dateEnd,
      location,
      durationHours,
      columnValues,
      attendees,
    };
  });
}

export type BoardColumn = { id: string; title: string; type: string };

export async function fetchBoardColumns(
  accessToken: string,
  boardId: bigint,
): Promise<BoardColumn[]> {
  const query = `
    query ($boardIds: [ID!]) {
      boards(ids: $boardIds) {
        columns {
          id
          title
          type
        }
      }
    }
  `;
  type Resp = {
    boards: Array<{
      columns: Array<{ id: string; title: string; type: string }>;
    }>;
  };
  const data = await queryMonday<Resp>(accessToken, query, { boardIds: [boardId.toString()] });
  return data.boards[0]?.columns ?? [];
}

export async function createMondayItem(
  accessToken: string,
  boardId: bigint,
  itemName: string,
  columnValues: Record<string, unknown>,
): Promise<{ id: string; updated_at: string }> {
  const query = `
    mutation ($boardId: ID!, $itemName: String!, $columnValues: JSON!) {
      create_item(board_id: $boardId, item_name: $itemName, column_values: $columnValues) {
        id
        updated_at
      }
    }
  `;
  type Resp = {
    create_item: { id: string; updated_at: string };
  };
  const data = await queryMonday<Resp>(accessToken, query, {
    boardId: boardId.toString(),
    itemName,
    columnValues: JSON.stringify(columnValues),
  });
  return data.create_item;
}

export async function updateMondayItem(
  accessToken: string,
  boardId: bigint,
  itemId: string,
  columnValues: Record<string, unknown>,
): Promise<{ id: string; updated_at: string }> {
  const query = `
    mutation ($boardId: ID!, $itemId: ID!, $columnValues: JSON!) {
      change_multiple_column_values(board_id: $boardId, item_id: $itemId, column_values: $columnValues) {
        id
        updated_at
      }
    }
  `;
  type Resp = {
    change_multiple_column_values: { id: string; updated_at: string };
  };
  const data = await queryMonday<Resp>(accessToken, query, {
    boardId: boardId.toString(),
    itemId,
    columnValues: JSON.stringify(columnValues),
  });
  return data.change_multiple_column_values;
}

export async function fetchMondayUsers(accessToken: string): Promise<MondayUser[]> {
  const query = `
    query {
      users {
        id
        name
        email
      }
    }
  `;
  type Resp = {
    users: Array<{ id: string; name: string; email: string }>;
  };
  const data = await queryMonday<Resp>(accessToken, query);
  return (data.users ?? []).map((u) => ({
    id: BigInt(u.id),
    name: u.name,
    email: u.email,
  }));
}

export function findPeopleColumn(columns: BoardColumn[]): BoardColumn | undefined {
  return (
    columns.find((c) => c.type === 'people') ??
    columns.find((c) => c.id.startsWith('multiple_person_'))
  );
}
