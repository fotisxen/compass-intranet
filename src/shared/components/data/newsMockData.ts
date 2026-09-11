export interface INewsItem {
  title: string;
  date: string;
}

export interface INewsCategory {
  label: string;
  highlighted?: boolean;
}

export const newsCategories: INewsCategory[] = [
  { label: 'All news' },
  { label: 'Corporate' },
  { label: 'Fleet' },
  { label: 'People & Culture' },
  { label: 'Open Positions', highlighted: true }
];

export const news: INewsItem[] = [
  { title: 'Taking to the slopes with Canadian photographer and ski Kari Medig', date: '14 September 2026' },
  { title: 'New wellness program launching in September', date: '19 August 2026' },
  { title: 'Athens office renovation complete', date: '14 August 2026' }
];
