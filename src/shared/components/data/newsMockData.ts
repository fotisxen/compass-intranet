export interface INewsItem {
  title: string;
  date: string;
  imageUrl: string;
}

export interface INewsCategory {
  label: string;
  highlighted?: boolean;
  /** Where the round arrow button next to the selected tab links to. Defaults to '#'. */
  href?: string;
}

export const newsCategories: INewsCategory[] = [
  { label: 'Corporate', href: '/sites/Intranet/SitePages/Internal-Company-Announcements.aspx#corporate-news' },
  { label: 'Fleet', href: '/sites/Intranet/SitePages/Internal-Company-Announcements.aspx#fleet-updates' },
  { label: 'People & Purpose', href: '/sites/Intranet/SitePages/Internal-Company-Announcements.aspx#people-culture-news' },
  { label: 'Organizational Development', href: '/sites/Intranet/SitePages/Organizational-Developments.aspx' },
  { label: 'Open Positions', href: '/sites/Intranet/SitePages/Career-Opportunities.aspx#open', highlighted: true }
];

export const news: INewsItem[] = [
  {
    title: 'Taking to the slopes with Canadian photographer and ski Kari Medig',
    date: '14 September 2026',
    imageUrl: 'https://images.unsplash.com/photo-1551524559-8af4e6624178?w=800&q=80'
  },
  {
    title: 'New wellness program launching in September',
    date: '19 August 2026',
    imageUrl: 'https://images.unsplash.com/photo-1544367567-0f2fcb009e0b?w=800&q=80'
  },
  {
    title: 'Athens office renovation complete',
    date: '14 August 2026',
    imageUrl: 'https://images.unsplash.com/photo-1497366216548-37526070297c?w=800&q=80'
  }
];
