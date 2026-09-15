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

export interface IStockTicker {
  symbol: string;
  changePct: string;
  price: string;
}

export const stockTickers: IStockTicker[] = [
  { symbol: 'SBLK', changePct: '+2,0% ↑', price: '$26.56' },
  { symbol: 'STK2', changePct: '+2,2% ↑', price: '$14.56' }
];
