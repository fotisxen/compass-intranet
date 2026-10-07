import { SPHttpClient } from '@microsoft/sp-http';
import { type INavItem, type INavLink } from './navData';
import { addItem, countItems, createList, ensureColumns, escapeXml, getList, listItemsUrl, type IColumnDef } from './spListSetup';

/**
 * The site navigation lives in a SharePoint list so owners can add, remove,
 * rename, reorder and re-point menu entries without a code change.
 *
 *  - A row with no "Parent menu item" is a top-level entry in the bar.
 *  - A row whose "Parent menu item" is another row is an entry in that
 *    entry's dropdown. Top-level entries with no children are plain links.
 *  - "Dropdown column" splits a long dropdown into side-by-side columns.
 *  - "Order" sorts entries (smaller first); "Show in menu" hides an entry
 *    without deleting it.
 */
export const NAV_LIST_TITLE = 'Compass Navigation';

// Empty SharePoint fields arrive as null at runtime; every use below tolerates that.
interface INavListItem {
  Id: number;
  Title?: string;
  NavUrl?: string;
  ParentItemId?: number;
  SortOrder?: number;
  DropdownColumn?: number;
  ShowInMenu?: boolean;
}

const NAV_COLUMNS: IColumnDef[] = [
  {
    name: 'NavUrl',
    schemaXml: () => '<Field Type="Text" Name="NavUrl" StaticName="NavUrl" DisplayName="Link URL" MaxLength="500" />'
  },
  {
    name: 'ParentItem',
    // Lookup back into this same list — a dropdown of the other menu entries.
    schemaXml: listId =>
      `<Field Type="Lookup" Name="ParentItem" StaticName="ParentItem" DisplayName="Parent menu item" List="{${escapeXml(listId)}}" ShowField="Title" Required="FALSE" />`
  },
  {
    name: 'SortOrder',
    schemaXml: () => '<Field Type="Number" Name="SortOrder" StaticName="SortOrder" DisplayName="Order" Decimals="0" />'
  },
  {
    name: 'DropdownColumn',
    schemaXml: () => '<Field Type="Number" Name="DropdownColumn" StaticName="DropdownColumn" DisplayName="Dropdown column" Decimals="0" Min="1" Max="6"><Default>1</Default></Field>'
  },
  {
    name: 'ShowInMenu',
    schemaXml: () => '<Field Type="Boolean" Name="ShowInMenu" StaticName="ShowInMenu" DisplayName="Show in menu"><Default>1</Default></Field>'
  }
];

function orderOf(item: INavListItem): number {
  return typeof item.SortOrder === 'number' ? item.SortOrder : 100000 + item.Id;
}

function byOrder(a: INavListItem, b: INavListItem): number {
  return orderOf(a) - orderOf(b) || a.Id - b.Id;
}

/** Turns the flat list rows into the nested structure the header renders. */
export function buildNavFromListItems(rows: INavListItem[]): INavItem[] {
  const visible = rows.filter(r => r.ShowInMenu !== false && !!r.Title && !!r.Title.trim());
  const children = (parentId: number): INavListItem[] => visible.filter(r => r.ParentItemId === parentId).sort(byOrder);

  return visible
    .filter(r => !r.ParentItemId)
    .sort(byOrder)
    .map((top): INavItem => {
      const sub = children(top.Id);
      if (sub.length === 0) {
        return { label: top.Title as string, href: top.NavUrl || '#' };
      }

      const columnNumbers = Array.from(new Set(sub.map(s => s.DropdownColumn || 1))).sort((a, b) => a - b);
      return {
        label: top.Title as string,
        columns: columnNumbers.map(n => [
          {
            items: sub
              .filter(s => (s.DropdownColumn || 1) === n)
              .map((s): INavLink => ({ label: s.Title as string, href: s.NavUrl || '#' }))
          }
        ])
      };
    });
}

export type NavListResult = INavItem[] | 'missing' | 'unreadable';

/**
 * Reads the navigation. 'missing' = the list doesn't exist, 'unreadable' =
 * it exists but can't be read as expected (e.g. columns not added yet).
 * Network failures reject, so callers can tell "no list" from "offline".
 */
export async function loadNavFromList(client: SPHttpClient, siteUrl: string, listTitle: string): Promise<NavListResult> {
  const select = 'Id,Title,NavUrl,ParentItemId,SortOrder,DropdownColumn,ShowInMenu';
  const response = await client.get(`${listItemsUrl(siteUrl, listTitle)}?$select=${select}&$top=500`, SPHttpClient.configurations.v1);
  if (response.status === 404) {
    return 'missing';
  }
  if (!response.ok) {
    return 'unreadable';
  }
  const data: { value: INavListItem[] } = await response.json();
  return buildNavFromListItems(data.value || []);
}

/**
 * Creates the list (with its columns) if needed and, when it is empty, fills
 * it with the given menu so the live menu doesn't change. Safe to re-run.
 */
export async function ensureNavigationList(client: SPHttpClient, siteUrl: string, listTitle: string, seed: INavItem[]): Promise<void> {
  const list = (await getList(client, siteUrl, listTitle))
    || (await createList(client, siteUrl, listTitle, 'Site navigation: top-level entries, their dropdown entries and where each one links to. Managed by the Compass header.'));

  await ensureColumns(client, siteUrl, listTitle, list.Id, NAV_COLUMNS);

  if ((await countItems(client, siteUrl, listTitle)) > 0) {
    return;
  }

  const entityType = list.ListItemEntityTypeFullName
    || (await getList(client, siteUrl, listTitle) || { ListItemEntityTypeFullName: '' }).ListItemEntityTypeFullName;

  for (let i = 0; i < seed.length; i++) {
    const top = seed[i];
    const parentId = await addItem(client, siteUrl, listTitle, entityType, {
      Title: top.label,
      NavUrl: top.href || '',
      SortOrder: (i + 1) * 10,
      DropdownColumn: 1,
      ShowInMenu: true
    });

    const columns = top.columns || [];
    for (let c = 0; c < columns.length; c++) {
      let position = 0;
      for (const group of columns[c]) {
        for (const link of group.items) {
          position += 1;
          await addItem(client, siteUrl, listTitle, entityType, {
            Title: link.label,
            NavUrl: link.href,
            ParentItemId: parentId,
            SortOrder: position * 10,
            DropdownColumn: c + 1,
            ShowInMenu: true
          });
        }
      }
    }
  }
}
