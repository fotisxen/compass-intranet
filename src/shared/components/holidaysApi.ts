import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import { type IHoliday } from './data/holidaysMockData';

// The real "Public Holidays" widget on the live site reads from a plain
// SharePoint Events list (Title + EventDate) — this is that list's GUID,
// found in the page's own web part configuration.
const PUBLIC_HOLIDAYS_LIST_ID = 'd6e5df27-ce8c-4770-9b39-e58c91c50d4b';
// Only the 2 nearest holidays (today or later) — the query already filters
// to >= today and sorts ascending, so this just caps the count.
const PAGE_SIZE = 2;

interface ISpEventItem {
  Title: string;
  EventDate: string;
}

interface ISpListItemsResponse {
  value: ISpEventItem[];
}

function formatShortDate(iso: string): string {
  return new Date(iso).toLocaleDateString('en-US', { month: 'short', day: 'numeric' }).toUpperCase();
}

/** Resolves to [] when the list can't be read, so callers just show nothing. */
export async function loadUpcomingHolidays(spHttpClient: SPHttpClient, siteUrl: string): Promise<IHoliday[]> {
  const today = new Date().toISOString();
  const endpoint =
    `${siteUrl}/_api/web/lists(guid'${PUBLIC_HOLIDAYS_LIST_ID}')/items` +
    `?$select=Title,EventDate&$filter=EventDate ge datetime'${today}'&$orderby=EventDate asc&$top=${PAGE_SIZE}`;

  const response: SPHttpClientResponse = await spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
  if (!response.ok) {
    return [];
  }

  const data: ISpListItemsResponse = await response.json();
  if (!data.value) {
    return [];
  }

  return data.value.map(item => ({ date: formatShortDate(item.EventDate), label: item.Title }));
}
