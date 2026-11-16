import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';

export interface INewsItem {
  title: string;
  date: string;
  url: string;
  imageUrl?: string;
}

interface IBannerImageField {
  Description?: string;
  Url?: string;
  serverRelativeUrl?: string;
}

interface ISpListItem {
  Title: string;
  FileRef: string;
  Created: string;
  // Site Pages' Banner Image is an "Image"-type column. Classic REST has been
  // observed returning this both as an already-parsed { Description, Url }
  // object and as a JSON-encoded string — handle both shapes.
  BannerImageUrl?: string | IBannerImageField;
}

interface ISpListItemsResponse {
  value: ISpListItem[];
}

const SELECT = '$select=Title,FileRef,Created,BannerImageUrl';

/**
 * SharePoint sometimes stores a banner/photo as a small generated preview
 * ("/_layouts/15/getpreview.ashx?path=<file>", only a few hundred px wide),
 * which looks blurry on a card. The file it points at is the full-size original,
 * so use that directly.
 */
export function originalImageUrl(url: string | undefined): string | undefined {
  if (!url) {
    return url;
  }
  const match = /^(https?:\/\/[^/]+)?\/_layouts\/15\/getpreview\.ashx\?(.*)$/i.exec(url.trim());
  if (!match) {
    return url;
  }
  const pathParam = /(?:^|&)path=([^&]+)/i.exec(match[2]);
  if (!pathParam) {
    return url;
  }
  try {
    return (match[1] || '') + encodeURI(decodeURIComponent(pathParam[1].replace(/\+/g, ' ')));
  } catch {
    return url;
  }
}

/**
 * Downloads the pictures into the browser cache ahead of time, so a card
 * appears (or a rotation swaps in) with its photo already there instead of the
 * text arriving first. Never fails and never waits longer than `maxWaitMs`.
 */
export function preloadImages(items: INewsItem[], maxWaitMs: number = 4000): Promise<void> {
  const loads = items
    .filter(i => !!i.imageUrl)
    .map(i => new Promise<void>(resolve => {
      const img = new Image();
      img.onload = () => resolve();
      img.onerror = () => resolve();
      img.src = i.imageUrl as string;
    }));
  return Promise.race([
    Promise.all(loads).then(() => undefined),
    new Promise<void>(resolve => window.setTimeout(resolve, maxWaitMs))
  ]);
}

export function parseBannerImageUrl(raw?: string | IBannerImageField): string | undefined {
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
      const parsed: IBannerImageField = JSON.parse(trimmed);
      return parsed.Url || parsed.serverRelativeUrl || parsed.Description || undefined;
    } catch {
      return undefined;
    }
  }
  return trimmed;
}

function toNewsItem(spItem: ISpListItem): INewsItem {
  return {
    title: spItem.Title,
    date: new Date(spItem.Created).toLocaleDateString('en-US', { day: 'numeric', month: 'long', year: 'numeric' }),
    url: spItem.FileRef,
    imageUrl: originalImageUrl(parseBannerImageUrl(spItem.BannerImageUrl))
  };
}

/** Accepts a full URL, a server-relative path, or just a page file name, and returns a lowercase server-relative path for comparing against FileRef. */
export function normalizePageUrl(siteUrl: string, raw: string): string {
  let value = raw.trim();
  if (!value) {
    return '';
  }
  const hashOrQuery = value.search(/[?#]/);
  if (hashOrQuery >= 0) {
    value = value.slice(0, hashOrQuery);
  }
  value = value.replace(/^https?:\/\/[^/]+/i, '');
  if (value.charAt(0) !== '/') {
    const webPath = siteUrl.replace(/^https?:\/\/[^/]+/i, '').replace(/\/$/, '');
    value = `${webPath}/SitePages/${value}`;
  }
  try {
    value = decodeURIComponent(value);
  } catch {
    // Leave it as typed if it isn't valid percent-encoding.
  }
  return value.toLowerCase();
}

/** Newest first. "Site Pages" with PromotedState eq 2 is how SharePoint marks a page as posted to News. Resolves to [] when the list can't be read. */
export async function loadPromotedNews(spHttpClient: SPHttpClient, siteUrl: string, top: number): Promise<INewsItem[]> {
  const endpoint =
    `${siteUrl}/_api/web/lists/GetByTitle('Site Pages')/items` +
    `?${SELECT}&$filter=PromotedState eq 2&$orderby=Created desc&$top=${top}`;

  const response: SPHttpClientResponse = await spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
  if (!response.ok) {
    return [];
  }
  const data: ISpListItemsResponse = await response.json();
  return (data.value || []).map(toNewsItem);
}

/** Any Site Pages page by its (normalized, server-relative) path — pinned news doesn't have to be promoted. */
export async function loadPageByPath(spHttpClient: SPHttpClient, siteUrl: string, serverRelativePath: string): Promise<INewsItem | undefined> {
  const endpoint =
    `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${serverRelativePath.replace(/'/g, "''")}')/ListItemAllFields` +
    `?${SELECT}`;

  const response: SPHttpClientResponse = await spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
  if (!response.ok) {
    return undefined;
  }
  const item: ISpListItem = await response.json();
  return item && item.Title ? toNewsItem(item) : undefined;
}
