import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * Shared by the Events Grid and Event Details/Image web parts: they all read the
 * site's SharePoint Events (calendar) list — the same one the home page's
 * Upcoming Events card uses.
 */
export const DEFAULT_EVENTS_LIST = 'dd3316a0-3bc0-4d4c-acff-71851299dad7';

export interface IEventItem {
  id: number;
  title: string;
  start: Date;
  /** Same as start when the item has no end date. */
  end: Date;
  allDay: boolean;
  location: string;
  category: string;
  /** The "Description" column: SharePoint rich text, sanitized by sanitizeHtml() before it is shown. */
  descriptionHtml: string;
  imageUrl?: string;
}

interface IThumbnailField {
  Description?: string;
  Url?: string;
  serverRelativeUrl?: string;
}

// Empty SharePoint fields arrive as null at runtime; every use below tolerates that.
interface IEventRow {
  Id: number;
  Title?: string;
  EventDate?: string;
  EndDate?: string;
  Location?: string;
  Description?: string;
  fAllDayEvent?: boolean;
  Category?: string;
  BannerUrl?: string | IThumbnailField;
}

// "Thumbnail"-type column: classic REST has been seen returning it both as a
// JSON-encoded blob and as an already-parsed object — handle both shapes.
export function parseThumbnailUrl(raw?: string | IThumbnailField): string | undefined {
  if (!raw) {
    return undefined;
  }
  if (typeof raw === 'object') {
    return raw.Url || raw.serverRelativeUrl || raw.Description || undefined;
  }
  const trimmed = raw.trim();
  if (!trimmed) {
    return undefined;
  }
  if (trimmed.charAt(0) === '{') {
    try {
      const parsed: IThumbnailField = JSON.parse(trimmed);
      return parsed.Url || parsed.serverRelativeUrl || parsed.Description || undefined;
    } catch {
      return undefined;
    }
  }
  return trimmed;
}

const WANTED_FIELDS = ['Id', 'Title', 'EventDate', 'EndDate', 'Location', 'Description', 'fAllDayEvent', 'Category', 'BannerUrl'];

function listApi(siteUrl: string, listRef: string): string {
  const ref = (listRef || DEFAULT_EVENTS_LIST).trim();
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(ref)
    ? `${siteUrl}/_api/web/lists(guid'${ref}')`
    : `${siteUrl}/_api/web/lists/getbytitle('${ref.replace(/'/g, "''")}')`;
}

// Which of the columns we would like exist on this particular list (a calendar
// without a Category or Banner column would otherwise fail the whole request).
const fieldCache: Record<string, Promise<string[]>> = {};

function availableFields(client: SPHttpClient, siteUrl: string, listRef: string): Promise<string[]> {
  const key = siteUrl + '|' + listRef;
  if (!fieldCache[key]) {
    fieldCache[key] = client.get(`${listApi(siteUrl, listRef)}/fields?$select=InternalName&$top=500`, SPHttpClient.configurations.v1)
      .then(async (response: SPHttpClientResponse) => {
        if (!response.ok) {
          return ['Id', 'Title', 'EventDate'];
        }
        const data: { value: { InternalName: string }[] } = await response.json();
        const names = (data.value || []).map(f => f.InternalName);
        return WANTED_FIELDS.filter(f => names.indexOf(f) !== -1 || f === 'Id');
      })
      .catch(() => ['Id', 'Title', 'EventDate']);
  }
  return fieldCache[key];
}

function toItem(row: IEventRow): IEventItem {
  const start = new Date(row.EventDate || Date.now());
  const end = row.EndDate ? new Date(row.EndDate) : start;
  return {
    id: row.Id,
    title: (row.Title || '').trim(),
    start,
    end: isNaN(end.getTime()) ? start : end,
    allDay: !!row.fAllDayEvent,
    location: (row.Location || '').trim(),
    category: (row.Category || '').trim(),
    descriptionHtml: row.Description || '',
    imageUrl: parseThumbnailUrl(row.BannerUrl)
  };
}

export interface ILoadEventsOptions {
  /** How many events to fetch (the soonest ones). */
  top: number;
}

export async function loadEvents(client: SPHttpClient, siteUrl: string, listRef: string, options: ILoadEventsOptions): Promise<IEventItem[]> {
  const fields = await availableFields(client, siteUrl, listRef);
  // Today and later only: an event counts by the day it starts.
  const startOfToday = new Date();
  startOfToday.setHours(0, 0, 0, 0);
  const filter = `&$filter=${encodeURIComponent(`EventDate ge datetime'${startOfToday.toISOString()}'`)}`;
  const url = `${listApi(siteUrl, listRef)}/items?$select=${fields.join(',')}${filter}&$orderby=EventDate asc&$top=${options.top}`;
  const response = await client.get(url, SPHttpClient.configurations.v1);
  if (!response.ok) {
    throw new Error(`Reading the events failed (${response.status})`);
  }
  const data: { value: IEventRow[] } = await response.json();
  // Checked again here, so an event from the past can never get through.
  return (data.value || []).filter(r => !!r.Title).map(toItem).filter(ev => ev.start.getTime() >= startOfToday.getTime());
}

/** One event by its list item id, or undefined when it does not exist. */
export async function loadEvent(client: SPHttpClient, siteUrl: string, listRef: string, id: number): Promise<IEventItem | undefined> {
  const fields = await availableFields(client, siteUrl, listRef);
  const response = await client.get(`${listApi(siteUrl, listRef)}/items(${id})?$select=${fields.join(',')}`, SPHttpClient.configurations.v1);
  if (!response.ok) {
    return undefined;
  }
  const row: IEventRow = await response.json();
  return row && row.Title ? toItem(row) : undefined;
}

/** The first upcoming event — shown while editing a detail page that has no ?id= yet. */
export async function loadNextEvent(client: SPHttpClient, siteUrl: string, listRef: string): Promise<IEventItem | undefined> {
  const events = await loadEvents(client, siteUrl, listRef, { top: 1 });
  return events[0];
}

/** The id a detail page was opened with: …/Event.aspx?eventId=12 (id= and ItemId= also work; SharePoint reserves id= on its own pages, so links use eventId). */
export function eventIdFromUrl(search: string = window.location.search): number | undefined {
  const params = new URLSearchParams(search);
  const raw = params.get('eventId') || params.get('id') || params.get('ItemId') || params.get('itemid');
  const id = raw ? parseInt(raw, 10) : NaN;
  return isNaN(id) ? undefined : id;
}

/** Link to the detail page for an event. */
export function eventUrl(pageUrl: string, id: number): string {
  return `${pageUrl}${pageUrl.indexOf('?') === -1 ? '?' : '&'}eventId=${id}`;
}

// ---------------------------------------------------------------- formatting

const EN = 'en-US';

/** "Fri, Oct 30, 10:00 AM" (or "Fri, Oct 30 · All day"). */
export function formatShort(event: IEventItem): string {
  const day = event.start.toLocaleDateString(EN, { weekday: 'short', month: 'short', day: 'numeric' });
  return event.allDay
    ? `${day} · All day`
    : `${day}, ${event.start.toLocaleTimeString(EN, { hour: 'numeric', minute: '2-digit' })}`;
}

/** "OCTOBER 2026" */
export function formatMonthHeading(date: Date): string {
  return `${date.toLocaleDateString(EN, { month: 'long' }).toUpperCase()} ${date.getFullYear()}`;
}

function longDate(date: Date, allDay: boolean): string {
  const day = date.toLocaleDateString(EN, { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
  return `${day}, ${allDay ? 'All day' : date.toLocaleTimeString(EN, { hour: 'numeric', minute: '2-digit' })}`;
}

/** The "When" lines of the detail page. */
export function formatWhen(event: IEventItem, showTimezone: boolean): string[] {
  const lines = [longDate(event.start, event.allDay)];
  if (event.end.getTime() !== event.start.getTime()) {
    lines.push(`- ${longDate(event.end, event.allDay)}`);
  }
  if (showTimezone) {
    const offsetMinutes = -event.start.getTimezoneOffset();
    const sign = offsetMinutes >= 0 ? '+' : '-';
    const abs = Math.abs(offsetMinutes);
    const offset = `${sign}${pad(Math.floor(abs / 60))}:${pad(abs % 60)}`;
    // "9:00:00 AM Eastern European Standard Time" -> keep only the zone's name.
    const zone = event.start.toLocaleTimeString(EN, { timeZoneName: 'long' }).replace(/^\d{1,2}:\d{2}:\d{2}\s*(AM|PM)?\s*/i, '').trim();
    const name = zone ? ` ${zone}` : '';
    lines.push(`(UTC${offset})${name}`);
  }
  return lines;
}

// ----------------------------------------------------------------- rich text

const DROPPED_TAGS = ['script', 'style', 'iframe', 'object', 'embed', 'link', 'meta', 'base', 'form', 'input', 'button', 'svg'];

/** Keeps SharePoint's rich text formatting but removes anything that could run code. */
export function sanitizeHtml(html: string): string {
  if (!html) {
    return '';
  }
  const doc = new DOMParser().parseFromString(html, 'text/html');
  DROPPED_TAGS.forEach(tag => {
    Array.prototype.forEach.call(doc.body.querySelectorAll(tag), (el: Element) => el.remove());
  });
  Array.prototype.forEach.call(doc.body.querySelectorAll('*'), (el: Element) => {
    Array.prototype.slice.call(el.attributes).forEach((attr: Attr) => {
      const name = attr.name.toLowerCase();
      const value = attr.value.trim().toLowerCase();
      // eslint-disable-next-line no-script-url
      if (name.indexOf('on') === 0 || ((name === 'href' || name === 'src' || name === 'xlink:href') && value.indexOf('javascript:') === 0)) {
        el.removeAttribute(attr.name);
      }
    });
    if (el.tagName === 'A') {
      el.setAttribute('rel', 'noopener noreferrer');
    }
  });
  return doc.body.innerHTML;
}

function plainText(html: string): string {
  return new DOMParser().parseFromString(html || '', 'text/html').body.textContent || '';
}

// ------------------------------------------------------------ add to calendar

function pad(n: number): string {
  return n < 10 ? `0${n}` : `${n}`;
}

function icsUtc(d: Date): string {
  return `${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`;
}

function icsDate(d: Date): string {
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
}

function icsText(value: string): string {
  return value.replace(/\\/g, '\\\\').replace(/\r?\n/g, '\\n').replace(/,/g, '\\,').replace(/;/g, '\\;');
}

/** The event as an iCalendar (.ics) file's text — opens in Outlook, Apple Calendar, Google Calendar. */
export function buildIcs(event: IEventItem, pageUrl: string): string {
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Compass//Events//EN', 'CALSCALE:GREGORIAN', 'BEGIN:VEVENT'];
  lines.push(`UID:event-${event.id}@compass-intranet`);
  lines.push(`DTSTAMP:${icsUtc(new Date())}`);
  if (event.allDay) {
    // An all-day event ends the day AFTER its last day in iCalendar.
    const lastDay = new Date(event.end.getFullYear(), event.end.getMonth(), event.end.getDate() + 1);
    lines.push(`DTSTART;VALUE=DATE:${icsDate(event.start)}`);
    lines.push(`DTEND;VALUE=DATE:${icsDate(lastDay)}`);
  } else {
    const end = event.end.getTime() > event.start.getTime() ? event.end : new Date(event.start.getTime() + 60 * 60 * 1000);
    lines.push(`DTSTART:${icsUtc(event.start)}`);
    lines.push(`DTEND:${icsUtc(end)}`);
  }
  lines.push(`SUMMARY:${icsText(event.title)}`);
  if (event.location) {
    lines.push(`LOCATION:${icsText(event.location)}`);
  }
  const text = plainText(event.descriptionHtml).trim();
  if (text) {
    lines.push(`DESCRIPTION:${icsText(text)}`);
  }
  lines.push(`URL:${pageUrl}`);
  lines.push('END:VEVENT', 'END:VCALENDAR');
  return lines.join('\r\n');
}

/** Saves the event as a .ics file (the browser downloads it; double-click adds it to the calendar). */
export function downloadIcs(event: IEventItem, pageUrl: string): void {
  const blob = new Blob([buildIcs(event, pageUrl)], { type: 'text/calendar;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${event.title.replace(/[^a-z0-9]+/gi, '-').replace(/^-|-$/g, '') || 'event'}.ics`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}
