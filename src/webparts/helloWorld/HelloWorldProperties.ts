import type { IAnniversary, IPersonSpotlight } from './components/data/mockData';

// Raw shapes stored in the web part's properties (edited via the property
// pane's PropertyFieldCollectionData panels). Kept flat (field1/2/3 instead
// of an arbitrary-length array) because that control needs a fixed set of
// columns per row.
export interface IPersonSpotlightItem {
  name: string;
  title: string;
  /** Only meaningful for the Promotions card — leave blank for Welcome Aboard. */
  newTitle?: string;
  /** Pasted image link (e.g. from Site Assets). Falls back to initials when blank or broken. */
  photoUrl?: string;
  field1Label: string;
  field1Value: string;
  field2Label: string;
  field2Value: string;
  field3Label: string;
  field3Value: string;
}

export interface IAnniversaryItem {
  name: string;
  role: string;
  years: number;
}

function initialsFor(name: string): string {
  const initials = (name || '')
    .trim()
    .split(/\s+/)
    .slice(0, 2)
    .map(w => w.charAt(0).toUpperCase())
    .join('');
  return initials || '?';
}

// Converts the property-pane's fixed field1/2/3 columns into the flexible
// {label, value}[] shape PersonSpotlightCard renders, dropping any pair the
// author left completely blank.
export function mapPersonItems(items: IPersonSpotlightItem[] | undefined): IPersonSpotlight[] {
  const source = items && items.length > 0 ? items : [];
  return source.map(item => ({
    name: item.name,
    title: item.title,
    newTitle: item.newTitle || undefined,
    photoUrl: item.photoUrl || undefined,
    initials: initialsFor(item.name),
    fields: [
      { label: item.field1Label, value: item.field1Value },
      { label: item.field2Label, value: item.field2Value },
      { label: item.field3Label, value: item.field3Value }
    ].filter(f => !!f.label && !!f.value)
  }));
}

export function mapAnniversaryItems(items: IAnniversaryItem[] | undefined): IAnniversary[] {
  const source = items && items.length > 0 ? items : [];
  return source.map(item => ({
    name: item.name,
    role: item.role,
    years: item.years,
    initials: initialsFor(item.name)
  }));
}

// Seeds a freshly-added web part (and any already-placed instance from
// before these properties existed) with the same content that used to be
// hard-coded in data/mockData.ts.
export const DEFAULT_WELCOME_ABOARD: IPersonSpotlightItem[] = [
  {
    name: 'Maria Papadopoulou',
    title: 'Crew Manager',
    field1Label: 'COMPANY',
    field1Value: 'Starbulk',
    field2Label: 'DEPARTMENT',
    field2Value: 'Fleet Operations',
    field3Label: 'REPORTS TO',
    field3Value: 'John Smith'
  }
];

export const DEFAULT_PROMOTIONS: IPersonSpotlightItem[] = [
  {
    name: 'Christos Markou',
    title: 'Deck Superintendent',
    newTitle: 'Fleet Operations Lead',
    field1Label: 'MOVE TYPE',
    field1Value: 'Promotion',
    field2Label: 'DEPARTMENT',
    field2Value: 'Fleet Operations',
    field3Label: 'REPORTS TO',
    field3Value: 'John Smith'
  }
];

export const DEFAULT_ANNIVERSARIES: IAnniversaryItem[] = [
  { name: 'Dimitris Perivolakis', role: 'HR Generalist', years: 1 },
  { name: 'Eleni Vasiliou', role: 'Chartering Analyst', years: 3 },
  { name: 'Nikos Andreou', role: 'Technical Superintendent', years: 10 },
  { name: 'Maria Papadopoulou', role: 'Crew Manager', years: 5 },
  { name: 'Giorgos Stefanidis', role: 'Operations Manager', years: 20 },
  { name: 'Anna Georgiou', role: 'Finance Controller', years: 1 }
];
