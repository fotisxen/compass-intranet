import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import { addItem, createList, ensureColumns, getList, listItemsUrl, type IColumnDef, type IListInfo } from '../components/spListSetup';

/**
 * The Marketplace's data lives in two SharePoint lists the app creates itself
 * (an owner opens a Marketplace web part once and presses the setup button):
 *
 *  - "Compass Marketplace": one row per listing. Photos are the row's list
 *    attachments. Content approval is switched on, so a new listing stays
 *    "Pending" (visible only to its author and the approvers) until it is
 *    approved; editing an approved listing sends the new version for approval
 *    again while the approved one stays public.
 *  - "Compass Marketplace Categories": the category chips. The site owners
 *    add, rename, reorder (SortOrder) or delete rows here; "All" is built in.
 */
export const LISTINGS_LIST_TITLE = 'Compass Marketplace';
export const CATEGORIES_LIST_TITLE = 'Compass Marketplace Categories';

/** Listings older than this many days stop showing (0 = never expire). */
export const DEFAULT_EXPIRY_DAYS = 30;
/** Most photos one listing can have. */
export const DEFAULT_MAX_PHOTOS = 5;

// SharePoint's built-in moderation status values.
export const STATUS_APPROVED = 0;
export const STATUS_REJECTED = 1;
export const STATUS_PENDING = 2;

export interface IListing {
  id: number;
  title: string;
  price: string;
  area: string;
  details: string;
  category: string;
  created: Date;
  authorName: string;
  authorEmail: string;
  /** Server-relative URLs, in upload order. */
  photos: string[];
  /** SharePoint moderation status (see STATUS_*). */
  status: number;
}

export interface ICategory {
  id: number;
  name: string;
}

export interface INewListing {
  title: string;
  price: string;
  area: string;
  details: string;
  category: string;
  /** Already resized (see imageResize.ts). */
  photos: { fileName: string; data: ArrayBuffer }[];
}

// Empty SharePoint fields arrive as null at runtime; every use below tolerates that.
interface IListingRow {
  Id: number;
  Title?: string;
  Price?: string;
  Area?: string;
  Details?: string;
  ListingCategory?: string;
  Created: string;
  OData__ModerationStatus?: number;
  Author?: { Title?: string; EMail?: string };
  AttachmentFiles?: { ServerRelativeUrl: string; FileName: string }[];
}

const SELECT = 'Id,Title,Price,Area,Details,ListingCategory,Created,OData__ModerationStatus,Author/Title,Author/EMail,AttachmentFiles/ServerRelativeUrl,AttachmentFiles/FileName';

const LISTING_COLUMNS: IColumnDef[] = [
  { name: 'Price', schemaXml: () => '<Field Type="Text" Name="Price" StaticName="Price" DisplayName="Price" MaxLength="100" />' },
  { name: 'Area', schemaXml: () => '<Field Type="Text" Name="Area" StaticName="Area" DisplayName="Area" MaxLength="100" />' },
  { name: 'ListingCategory', schemaXml: () => '<Field Type="Text" Name="ListingCategory" StaticName="ListingCategory" DisplayName="Category" MaxLength="100" />' },
  { name: 'Details', schemaXml: () => '<Field Type="Note" Name="Details" StaticName="Details" DisplayName="Details" NumLines="6" RichText="FALSE" />' }
];

const CATEGORY_COLUMNS: IColumnDef[] = [
  { name: 'SortOrder', schemaXml: () => '<Field Type="Number" Name="SortOrder" StaticName="SortOrder" DisplayName="Sort order" Decimals="0" />' }
];

const DEFAULT_CATEGORIES = ['For sale', 'Free', 'Wanted', 'Swap', 'Housing', 'Carpool', 'Lost & found'];

const JSON_HEADERS = { Accept: 'application/json;odata=nometadata' };
const WRITE_HEADERS = { ...JSON_HEADERS, 'Content-Type': 'application/json;odata=verbose', 'odata-version': '' };

function listApi(siteUrl: string, listTitle: string): string {
  return `${siteUrl}/_api/web/lists/getbytitle('${listTitle.replace(/'/g, "''")}')`;
}

async function failIfNotOk(response: SPHttpClientResponse, what: string): Promise<void> {
  if (!response.ok) {
    let detail = '';
    try {
      detail = await response.text();
    } catch {
      // No body to report.
    }
    throw new Error(`${what} failed (${response.status}) ${detail.slice(0, 200)}`);
  }
}

function toListing(row: IListingRow): IListing {
  return {
    id: row.Id,
    title: (row.Title || '').trim(),
    price: (row.Price || '').trim(),
    area: (row.Area || '').trim(),
    details: (row.Details || '').trim(),
    category: (row.ListingCategory || '').trim(),
    created: new Date(row.Created),
    authorName: (row.Author && row.Author.Title) || '',
    authorEmail: (row.Author && row.Author.EMail) || '',
    photos: (row.AttachmentFiles || []).map(f => f.ServerRelativeUrl),
    status: typeof row.OData__ModerationStatus === 'number' ? row.OData__ModerationStatus : STATUS_APPROVED
  };
}

function expiryFilter(expiryDays: number): string {
  if (!expiryDays || expiryDays <= 0) {
    return '';
  }
  const since = new Date(Date.now() - expiryDays * 24 * 60 * 60 * 1000).toISOString();
  return `Created ge datetime'${since}'`;
}

export type LoadResult<T> = { data: T } | 'missing' | 'unreadable';

/** The approved, not-yet-expired listings, newest first. Everyone — approvers included — only sees approved ones here. */
export async function loadListings(client: SPHttpClient, siteUrl: string, expiryDays: number): Promise<LoadResult<IListing[]>> {
  const filters = ['OData__ModerationStatus eq 0', expiryFilter(expiryDays)].filter(f => !!f).join(' and ');
  const url = `${listItemsUrl(siteUrl, LISTINGS_LIST_TITLE)}?$select=${SELECT}&$expand=Author,AttachmentFiles&$filter=${encodeURIComponent(filters)}&$orderby=Created desc&$top=200`;
  const response = await client.get(url, SPHttpClient.configurations.v1);
  if (response.status === 404) {
    return 'missing';
  }
  if (!response.ok) {
    return 'unreadable';
  }
  const data: { value: IListingRow[] } = await response.json();
  return { data: (data.value || []).filter(r => !!r.Title).map(toListing) };
}

async function getCurrentUserId(client: SPHttpClient, siteUrl: string): Promise<number | undefined> {
  const response = await client.get(`${siteUrl}/_api/web/currentuser?$select=Id`, SPHttpClient.configurations.v1);
  if (!response.ok) {
    return undefined;
  }
  const user: { Id: number } = await response.json();
  return user.Id;
}

/** The signed-in user's own listings in any status (they can see their own pending ones), newest first. */
export async function loadMyListings(client: SPHttpClient, siteUrl: string, top: number): Promise<LoadResult<IListing[]>> {
  const userId = await getCurrentUserId(client, siteUrl);
  if (userId === undefined) {
    return 'unreadable';
  }
  const url = `${listItemsUrl(siteUrl, LISTINGS_LIST_TITLE)}?$select=${SELECT}&$expand=Author,AttachmentFiles&$filter=${encodeURIComponent(`AuthorId eq ${userId}`)}&$orderby=Created desc&$top=${top}`;
  const response = await client.get(url, SPHttpClient.configurations.v1);
  if (response.status === 404) {
    return 'missing';
  }
  if (!response.ok) {
    return 'unreadable';
  }
  const data: { value: IListingRow[] } = await response.json();
  return { data: (data.value || []).filter(r => !!r.Title).map(toListing) };
}

/** Category names in the order the owners set (SortOrder, then name). */
export async function loadCategories(client: SPHttpClient, siteUrl: string): Promise<LoadResult<ICategory[]>> {
  const url = `${listItemsUrl(siteUrl, CATEGORIES_LIST_TITLE)}?$select=Id,Title,SortOrder&$top=100`;
  const response = await client.get(url, SPHttpClient.configurations.v1);
  if (response.status === 404) {
    return 'missing';
  }
  if (!response.ok) {
    return 'unreadable';
  }
  const data: { value: { Id: number; Title?: string; SortOrder?: number | null }[] } = await response.json();
  const rows = (data.value || []).filter(r => !!r.Title && !!r.Title.trim());
  rows.sort((a, b) => ((a.SortOrder ?? 9999) - (b.SortOrder ?? 9999)) || (a.Title || '').localeCompare(b.Title || ''));
  return { data: rows.map(r => ({ id: r.Id, name: (r.Title || '').trim() })) };
}

/** Turns on content approval and "edit only your own items" for the listings list. */
async function configureListingsList(client: SPHttpClient, siteUrl: string): Promise<void> {
  const response = await client.post(listApi(siteUrl, LISTINGS_LIST_TITLE), SPHttpClient.configurations.v1, {
    headers: { ...WRITE_HEADERS, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
    body: JSON.stringify({
      __metadata: { type: 'SP.List' },
      EnableVersioning: true,
      EnableModeration: true,
      // 2 = drafts (pending items) visible only to approvers and their author.
      DraftVersionVisibility: 2,
      // 2 = people can create items and edit only the ones they created.
      WriteSecurity: 2,
      ReadSecurity: 1
    })
  });
  await failIfNotOk(response, 'Setting up approval on the Marketplace list');
}

/** Creates both lists (columns, approval settings, default categories). Safe to re-run. Needs "manage lists" permission. */
export async function ensureMarketplaceLists(client: SPHttpClient, siteUrl: string): Promise<void> {
  let listings = await getList(client, siteUrl, LISTINGS_LIST_TITLE);
  if (!listings) {
    listings = await createList(client, siteUrl, LISTINGS_LIST_TITLE, 'Items employees post in the Compass Marketplace. New items wait for approval before they show.');
  }
  await ensureColumns(client, siteUrl, LISTINGS_LIST_TITLE, listings.Id, LISTING_COLUMNS);
  await configureListingsList(client, siteUrl);

  let categories = await getList(client, siteUrl, CATEGORIES_LIST_TITLE);
  const categoriesAreNew = !categories;
  if (!categories) {
    categories = await createList(client, siteUrl, CATEGORIES_LIST_TITLE, 'The category buttons of the Compass Marketplace. Add, rename, reorder (Sort order) or delete rows freely.');
  }
  await ensureColumns(client, siteUrl, CATEGORIES_LIST_TITLE, categories.Id, CATEGORY_COLUMNS);
  if (categoriesAreNew) {
    const info = (await getList(client, siteUrl, CATEGORIES_LIST_TITLE)) || categories;
    for (let i = 0; i < DEFAULT_CATEGORIES.length; i++) {
      await addItem(client, siteUrl, CATEGORIES_LIST_TITLE, info.ListItemEntityTypeFullName, { Title: DEFAULT_CATEGORIES[i], SortOrder: i + 1 });
    }
  }
}

/** Creates the listing (it starts as "Pending") and uploads its photos. Resolves to how many photos could not be uploaded. */
export async function submitListing(client: SPHttpClient, siteUrl: string, listing: INewListing): Promise<{ failedPhotos: number }> {
  const info: IListInfo | undefined = await getList(client, siteUrl, LISTINGS_LIST_TITLE);
  if (!info) {
    throw new Error('The Marketplace list does not exist yet.');
  }
  const id = await addItem(client, siteUrl, LISTINGS_LIST_TITLE, info.ListItemEntityTypeFullName, {
    Title: listing.title,
    Price: listing.price,
    Area: listing.area,
    Details: listing.details,
    ListingCategory: listing.category
  });

  let failedPhotos = 0;
  for (const photo of listing.photos) {
    const url = `${listApi(siteUrl, LISTINGS_LIST_TITLE)}/items(${id})/AttachmentFiles/add(FileName='${encodeURIComponent(photo.fileName)}')`;
    try {
      const response = await client.post(url, SPHttpClient.configurations.v1, { headers: JSON_HEADERS, body: photo.data });
      if (!response.ok) {
        failedPhotos++;
      }
    } catch {
      failedPhotos++;
    }
  }
  return { failedPhotos };
}

/** "Giorgos Tsimpoukelis" -> "Giorgos T." */
export function shortName(fullName: string): string {
  const parts = fullName.trim().split(/\s+/).filter(p => !!p);
  if (parts.length === 0) {
    return '';
  }
  if (parts.length === 1) {
    return parts[0];
  }
  return `${parts[0]} ${parts[parts.length - 1].charAt(0).toUpperCase()}.`;
}

export function initialsOf(fullName: string): string {
  const parts = fullName.trim().split(/\s+/).filter(p => !!p);
  return ((parts[0] || '').charAt(0) + (parts.length > 1 ? parts[parts.length - 1].charAt(0) : '')).toUpperCase() || '?';
}

/** "3 days ago" style age. */
export function timeAgo(date: Date, now: number = Date.now()): string {
  const minutes = Math.max(0, Math.round((now - date.getTime()) / 60000));
  if (minutes < 60) {
    return minutes <= 1 ? 'just now' : `${minutes} min ago`;
  }
  const hours = Math.round(minutes / 60);
  if (hours < 24) {
    return hours === 1 ? '1 hour ago' : `${hours} hours ago`;
  }
  const days = Math.round(hours / 24);
  if (days < 30) {
    return days === 1 ? '1 day ago' : `${days} days ago`;
  }
  const months = Math.round(days / 30);
  return months <= 1 ? '1 month ago' : `${months} months ago`;
}

/** Opens a Teams chat with this person (works from the browser and the desktop app). */
export function teamsChatUrl(email: string): string {
  return `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(email)}`;
}
