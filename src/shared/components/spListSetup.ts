import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';

// Small helpers for creating and filling a SharePoint list over REST. They
// exist so a site owner never has to build the "Compass Navigation" /
// "Compass Fleet Vessels" lists by hand — the app creates them (with the
// right columns) the first time an owner opens a page that needs them.

const WRITE_HEADERS = {
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=verbose',
  'odata-version': ''
};

export interface IListInfo {
  Id: string;
  ListItemEntityTypeFullName: string;
  /** Server-relative URL of the list's default view, for linking owners to it. */
  DefaultViewUrl?: string;
}

export interface IColumnDef {
  /** Internal name — what the code reads and writes. */
  name: string;
  /** Builds the field's schema XML; the list's id is passed for lookup columns that point back at the list itself. */
  schemaXml: (listId: string) => string;
}

function escapeForUrl(value: string): string {
  return value.replace(/'/g, "''");
}

export function escapeXml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function listUrl(siteUrl: string, listTitle: string): string {
  return `${siteUrl}/_api/web/lists/getbytitle('${escapeForUrl(listTitle)}')`;
}

async function failIfNotOk(response: SPHttpClientResponse, what: string): Promise<void> {
  if (!response.ok) {
    let detail = '';
    try {
      detail = await response.text();
    } catch {
      // No body to report.
    }
    throw new Error(`${what} failed (${response.status}) ${detail.slice(0, 300)}`);
  }
}

/** Resolves to undefined when the list doesn't exist (404). */
export async function getList(client: SPHttpClient, siteUrl: string, listTitle: string): Promise<IListInfo | undefined> {
  const response = await client.get(`${listUrl(siteUrl, listTitle)}?$select=Id,ListItemEntityTypeFullName,DefaultViewUrl`, SPHttpClient.configurations.v1);
  if (response.status === 404) {
    return undefined;
  }
  await failIfNotOk(response, 'Reading the list');
  return response.json();
}

export async function createList(client: SPHttpClient, siteUrl: string, listTitle: string, description: string): Promise<IListInfo> {
  const response = await client.post(`${siteUrl}/_api/web/lists`, SPHttpClient.configurations.v1, {
    headers: WRITE_HEADERS,
    body: JSON.stringify({
      __metadata: { type: 'SP.List' },
      BaseTemplate: 100,
      Title: listTitle,
      Description: description
    })
  });
  await failIfNotOk(response, 'Creating the list');
  const created: { Id: string } = await response.json();
  const info = await getList(client, siteUrl, listTitle);
  return info || { Id: created.Id, ListItemEntityTypeFullName: '' };
}

async function getFieldNames(client: SPHttpClient, siteUrl: string, listTitle: string): Promise<string[]> {
  const response = await client.get(`${listUrl(siteUrl, listTitle)}/fields?$select=InternalName`, SPHttpClient.configurations.v1);
  await failIfNotOk(response, 'Reading the list columns');
  const data: { value: { InternalName: string }[] } = await response.json();
  return data.value.map(f => f.InternalName);
}

/** Adds whichever of the columns are missing, so it is safe to run repeatedly. */
export async function ensureColumns(client: SPHttpClient, siteUrl: string, listTitle: string, listId: string, columns: IColumnDef[]): Promise<void> {
  const existing = await getFieldNames(client, siteUrl, listTitle);
  for (const column of columns) {
    if (existing.indexOf(column.name) !== -1) {
      continue;
    }
    const response = await client.post(`${listUrl(siteUrl, listTitle)}/fields/createfieldasxml`, SPHttpClient.configurations.v1, {
      headers: WRITE_HEADERS,
      body: JSON.stringify({
        parameters: {
          __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' },
          SchemaXml: column.schemaXml(listId),
          // 1 = also show it in the list's default view, 8 = keep the internal name as given.
          Options: 9
        }
      })
    });
    await failIfNotOk(response, `Adding the "${column.name}" column`);
  }
}

export async function countItems(client: SPHttpClient, siteUrl: string, listTitle: string): Promise<number> {
  const response = await client.get(`${listUrl(siteUrl, listTitle)}/ItemCount`, SPHttpClient.configurations.v1);
  await failIfNotOk(response, 'Counting the list items');
  const data: { value: number } = await response.json();
  return data.value;
}

/** Returns the new item's id. */
export async function addItem(client: SPHttpClient, siteUrl: string, listTitle: string, entityType: string, values: Record<string, unknown>): Promise<number> {
  const response = await client.post(`${listUrl(siteUrl, listTitle)}/items`, SPHttpClient.configurations.v1, {
    headers: WRITE_HEADERS,
    body: JSON.stringify({ __metadata: { type: entityType }, ...values })
  });
  await failIfNotOk(response, 'Adding a list item');
  const created: { Id: number } = await response.json();
  return created.Id;
}

/** Shows the built-in Title column under a friendlier name (e.g. "Vessel name"). Cosmetic, so failures are ignored. */
export async function renameTitleColumn(client: SPHttpClient, siteUrl: string, listTitle: string, displayName: string): Promise<void> {
  try {
    await client.post(`${listUrl(siteUrl, listTitle)}/fields/getbytitle('Title')`, SPHttpClient.configurations.v1, {
      headers: { ...WRITE_HEADERS, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
      body: JSON.stringify({ __metadata: { type: 'SP.Field' }, Title: displayName })
    });
  } catch {
    // Keeping the default name "Title" is fine.
  }
}

export function listItemsUrl(siteUrl: string, listTitle: string): string {
  return `${listUrl(siteUrl, listTitle)}/items`;
}
