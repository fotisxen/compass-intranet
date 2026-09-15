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
