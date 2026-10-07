import { SPHttpClient } from '@microsoft/sp-http';
import { createList, ensureColumns, getList, listItemsUrl, renameTitleColumn, type IColumnDef, type IListInfo } from './spListSetup';

/**
 * The fleet table's data: one row per vessel in a SharePoint list, so the
 * fleet can be edited (or pasted in from Excel via "Edit in grid view")
 * without touching the web part.
 */
export const FLEET_LIST_TITLE = 'Compass Fleet Vessels';

export interface IVessel {
  id: number;
  name: string;
  type: string;
  dwt?: number;
  built: string;
  shipyard: string;
}

// Empty SharePoint fields arrive as null at runtime; every use below tolerates that.
interface IVesselListItem {
  Id: number;
  Title?: string;
  VesselType?: string;
  Dwt?: number;
  BuiltYear?: string;
  Shipyard?: string;
}

const FLEET_COLUMNS: IColumnDef[] = [
  {
    name: 'VesselType',
    schemaXml: () => '<Field Type="Text" Name="VesselType" StaticName="VesselType" DisplayName="Vessel type" MaxLength="100" />'
  },
  {
    name: 'Dwt',
    schemaXml: () => '<Field Type="Number" Name="Dwt" StaticName="Dwt" DisplayName="DWT" Decimals="0" Min="0" />'
  },
  {
    // Text, not Number: a Number column would display a year as "2,015".
    name: 'BuiltYear',
    schemaXml: () => '<Field Type="Text" Name="BuiltYear" StaticName="BuiltYear" DisplayName="Built (year)" MaxLength="10" />'
  },
  {
    name: 'Shipyard',
    schemaXml: () => '<Field Type="Text" Name="Shipyard" StaticName="Shipyard" DisplayName="Shipyard" MaxLength="150" />'
  }
];

export type FleetListResult = { vessels: IVessel[] } | 'missing' | 'unreadable';

/** 'missing' = no such list, 'unreadable' = exists but can't be read as expected (e.g. columns not added yet). */
export async function loadVessels(client: SPHttpClient, siteUrl: string, listTitle: string): Promise<FleetListResult> {
  const select = 'Id,Title,VesselType,Dwt,BuiltYear,Shipyard';
  const response = await client.get(`${listItemsUrl(siteUrl, listTitle)}?$select=${select}&$top=500`, SPHttpClient.configurations.v1);
  if (response.status === 404) {
    return 'missing';
  }
  if (!response.ok) {
    return 'unreadable';
  }
  const data: { value: IVesselListItem[] } = await response.json();
  return {
    vessels: (data.value || [])
      .filter(r => !!r.Title && !!r.Title.trim())
      .map(r => ({
        id: r.Id,
        name: (r.Title || '').trim(),
        type: (r.VesselType || '').trim(),
        dwt: typeof r.Dwt === 'number' ? r.Dwt : undefined,
        built: (r.BuiltYear || '').trim(),
        shipyard: (r.Shipyard || '').trim()
      }))
  };
}

/** Creates the (empty) list with its columns. Safe to re-run. */
export async function ensureFleetList(client: SPHttpClient, siteUrl: string, listTitle: string): Promise<IListInfo> {
  let list = await getList(client, siteUrl, listTitle);
  const isNew = !list;
  if (!list) {
    list = await createList(client, siteUrl, listTitle, 'The vessels shown by the Compass Fleet Table web part. One row per vessel.');
  }
  await ensureColumns(client, siteUrl, listTitle, list.Id, FLEET_COLUMNS);
  if (isNew) {
    await renameTitleColumn(client, siteUrl, listTitle, 'Vessel name');
  }
  return (await getList(client, siteUrl, listTitle)) || list;
}

export type VesselSortField = 'dwt' | 'name' | 'built';

export function sortVessels(vessels: IVessel[], field: VesselSortField, descending: boolean): IVessel[] {
  const direction = descending ? -1 : 1;
  const compare = (a: IVessel, b: IVessel): number => {
    if (field === 'name') {
      return a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });
    }
    if (field === 'built') {
      return (parseInt(a.built, 10) || 0) - (parseInt(b.built, 10) || 0);
    }
    return (a.dwt || 0) - (b.dwt || 0);
  };
  // Rows missing the sort value always go last, whichever way it's sorted.
  const hasValue = (v: IVessel): boolean => (field === 'dwt' ? v.dwt !== undefined : field === 'built' ? !!parseInt(v.built, 10) : true);
  return vessels.slice().sort((a, b) => {
    if (hasValue(a) !== hasValue(b)) {
      return hasValue(a) ? -1 : 1;
    }
    return direction * compare(a, b) || a.name.localeCompare(b.name);
  });
}

/** 209537 -> "209.537" (dot as the thousands separator, as in the reference design). */
export function formatDwt(value: number): string {
  return String(Math.round(value)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}
