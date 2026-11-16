import { SPHttpClient } from '@microsoft/sp-http';
import { downscaleImage } from '../../shared/components/downscaleImage';

// Where the employee pictures live. A person with no Photo URL typed in gets
// the picture in this folder whose file name matches theirs.
const PHOTO_FOLDER = 'Photo Gallery/Employees Photos';

// The circles are 90px; this is plenty for a 2x screen.
const PERSON_PHOTO_WIDTH = 300;

export type PhotoFinder = (name: string) => string | undefined;

interface IPhotoFile {
  Name: string;
  ServerRelativeUrl: string;
}

function words(text: string): string[] {
  return text
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/\.[a-z0-9]{2,4}$/, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .split(' ')
    .filter(w => !!w);
}

/** Loads the file list once and returns a function that finds a person's picture by name ("Eva Michaloutsou" matches "E. Michaloutsou.png"). */
export async function loadPhotoFinder(spHttpClient: SPHttpClient, siteUrl: string): Promise<PhotoFinder> {
  const webPath = siteUrl.replace(/^https?:\/\/[^/]+/i, '').replace(/\/$/, '');
  const origin = (/^https?:\/\/[^/]+/i.exec(siteUrl) || [''])[0];
  const folder = `${webPath}/${PHOTO_FOLDER}`.replace(/'/g, "''");
  const endpoint = `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${folder}')/Files?$select=Name,ServerRelativeUrl&$top=5000`;
  const response = await spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
  if (!response.ok) {
    return () => undefined;
  }
  const data: { value?: IPhotoFile[] } = await response.json();
  const files = (data.value || [])
    .filter(f => /\.(jpe?g|png|webp)$/i.test(f.Name))
    .map(f => ({ tokens: words(f.Name), url: origin + encodeURI(f.ServerRelativeUrl) }));

  return (name: string): string | undefined => {
    const parts = words(name);
    if (parts.length < 2) {
      return undefined;
    }
    const first = parts[0];
    const last = parts[parts.length - 1];
    // Whole name first, then "first initial + surname".
    const exact = files.filter(f => f.tokens.indexOf(first) >= 0 && f.tokens.indexOf(last) >= 0)[0];
    const initial = files.filter(f => f.tokens.indexOf(last) >= 0 && f.tokens.some(t => t !== last && t.charAt(0) === first.charAt(0) && t.length <= 2))[0];
    const found = exact || initial;
    return found ? found.url : undefined;
  };
}

/** Downloads the pictures ahead of time, so the cards draw with their photos in place. Never fails and never waits longer than maxWaitMs. */
export function preloadPhotos(urls: (string | undefined)[], maxWaitMs: number = 4000): Promise<void> {
  const loads = urls
    .filter((u): u is string => !!u)
    // Photos are shown as small circles, so a shrunk copy is both sharper and quicker.
    .map(u => downscaleImage(u, PERSON_PHOTO_WIDTH).then(() => undefined));
  return Promise.race([
    Promise.all(loads).then(() => undefined),
    new Promise<void>(resolve => window.setTimeout(resolve, maxWaitMs))
  ]);
}
