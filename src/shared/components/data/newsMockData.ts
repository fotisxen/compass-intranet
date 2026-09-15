export interface INewsItem {
  title: string;
  date: string;
  imageUrl: string;
  /** Must match one of the real NewsCategory field's Choices exactly (see NewsCard's NEWS_CATEGORY_FIELD_NAME) — confirmed via _api/web/lists/GetByTitle('Site Pages')/fields?$filter=InternalName eq 'NewsCategory': Fleet & Operations, Sustainability, Safety, Community, Corporate. */
  category: string;
}

export interface INewsCategory {
  label: string;
  highlighted?: boolean;
  /** Where the round arrow button next to the selected tab links to. Defaults to '#'. */
  href?: string;
}

export const newsCategories: INewsCategory[] = [
  { label: 'Corporate', href: '/sites/Intranet/SitePages/Internal-Company-Announcements.aspx#corporate-news' },
  { label: 'Fleet & Operations', href: '/sites/Intranet/SitePages/Internal-Company-Announcements.aspx#fleet-updates' },
  { label: 'Community', href: '/sites/Intranet/SitePages/Internal-Company-Announcements.aspx#people-culture-news' },
  { label: 'Sustainability', href: '/sites/Intranet/SitePages/Internal-Company-Announcements.aspx' },
  { label: 'Safety', href: '/sites/Intranet/SitePages/Internal-Company-Announcements.aspx' }
];

export const news: INewsItem[] = [
  {
    title: 'Athens office renovation complete',
    date: '14 August 2026',
    imageUrl: 'https://images.unsplash.com/photo-1497366216548-37526070297c?w=800&q=80',
    category: 'Corporate'
  },
  {
    title: 'Piraeus terminal throughput hits a new fleet-handling record',
    date: '2 August 2026',
    imageUrl: 'https://images.unsplash.com/photo-1494412651409-8963ce7935a7?w=800&q=80',
    category: 'Fleet & Operations'
  },
  {
    title: 'New wellness program launching in September',
    date: '19 August 2026',
    imageUrl: 'https://images.unsplash.com/photo-1544367567-0f2fcb009e0b?w=800&q=80',
    category: 'Community'
  },
  {
    title: 'Fleet emissions down 8% year-on-year after retrofit program',
    date: '9 August 2026',
    imageUrl: 'https://images.unsplash.com/photo-1466611653911-95081537e5b7?w=800&q=80',
    category: 'Sustainability'
  },
  {
    title: 'Updated onboard safety drill schedule for Q4',
    date: '4 August 2026',
    imageUrl: 'https://images.unsplash.com/photo-1504328345606-18bbc8c9d7d1?w=800&q=80',
    category: 'Safety'
  }
];
