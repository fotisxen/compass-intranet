export interface IHoliday {
  date: string;
  label: string;
}

export interface IEvent {
  date: string;
  title: string;
  initials: string;
}

export interface IPersonSpotlight {
  name: string;
  title: string;
  /** When set, `title` renders as the previous (pale) role and `newTitle` as the new (bold) one. */
  newTitle?: string;
  initials: string;
  fields: { label: string; value: string }[];
}

export interface IAnniversary {
  name: string;
  role: string;
  years: number;
  initials: string;
}

export interface IOffice {
  city: string;
  country: string;
  lat: number;
  lng: number;
  address: string;
}

export const holidays: IHoliday[] = [
  { date: 'AUG 15', label: 'GR offices closed' },
  { date: 'OCT 28', label: 'GR offices closed' }
];

export const events: IEvent[] = [
  { date: '14.09.2026', title: 'Marketing Workshop', initials: 'MW' }
];

export const welcomeAboard: IPersonSpotlight = {
  name: 'Maria Papadopoulou',
  title: 'Crew Manager',
  initials: 'MP',
  fields: [
    { label: 'COMPANY', value: 'Starbulk' },
    { label: 'DEPARTMENT', value: 'Fleet Operations' },
    { label: 'REPORTS TO', value: 'John Smith' }
  ]
};

export const promotion: IPersonSpotlight = {
  name: 'Christos Markou',
  title: 'Deck Superintendent',
  newTitle: 'Fleet Operations Lead',
  initials: 'CM',
  fields: [
    { label: 'MOVE TYPE', value: 'Promotion' },
    { label: 'DEPARTMENT', value: 'Fleet Operations' },
    { label: 'REPORTS TO', value: 'John Smith' }
  ]
};

export const anniversaries: IAnniversary[] = [
  { name: 'Dimitris Perivolakis', role: 'HR Generalist', years: 1, initials: 'DP' },
  { name: 'Eleni Vasiliou', role: 'Chartering Analyst', years: 3, initials: 'EV' },
  { name: 'Nikos Andreou', role: 'Technical Superintendent', years: 10, initials: 'NA' },
  { name: 'Maria Papadopoulou', role: 'Crew Manager', years: 5, initials: 'MP' },
  { name: 'Giorgos Stefanidis', role: 'Operations Manager', years: 20, initials: 'GS' },
  { name: 'Anna Georgiou', role: 'Finance Controller', years: 1, initials: 'AG' }
];

export const offices: IOffice[] = [
  { city: 'Athens', country: 'Greece', lat: 37.9838, lng: 23.7275, address: 'HQ — Vasilissis Sofias Ave 1' },
  { city: 'London', country: 'United Kingdom', lat: 51.5072, lng: -0.1276, address: '10 Canada Square' },
  { city: 'New York', country: 'United States', lat: 40.7128, lng: -74.006, address: '350 5th Ave' },
  { city: 'Singapore', country: 'Singapore', lat: 1.3521, lng: 103.8198, address: '1 Raffles Place' },
  { city: 'Sydney', country: 'Australia', lat: -33.8688, lng: 151.2093, address: '1 Macquarie Pl' }
];

export interface IStockTicker {
  symbol: string;
  changePct: string;
  price: string;
}

export const stockTickers: IStockTicker[] = [
  { symbol: 'SBLK', changePct: '+2,0% ↑', price: '$26.56' },
  { symbol: 'STK2', changePct: '+2,2% ↑', price: '$14.56' }
];
