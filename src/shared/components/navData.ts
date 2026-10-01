export interface INavLink {
  label: string;
  href: string;
}

export interface INavGroup {
  heading?: string;
  headingHref?: string;
  items: INavLink[];
}

export type INavColumn = INavGroup[];

export interface INavItem {
  label: string;
  href?: string;
  columns?: INavColumn[];
}

export const NAV_ITEMS: INavItem[] = [
  {
    label: 'People',
    columns: [
      [
        {
          items: [
            { label: 'Onboarding', href: '/sites/Intranet/SitePages/Onboarding---Welcome-Aboard!.aspx' },
            { label: 'Organizational Developments', href: '/sites/Intranet/SitePages/Organizational-Developments.aspx' },
            { label: 'Open Positions', href: '/sites/Intranet/SitePages/Career-Opportunities.aspx' },
            { label: 'Blood Donation', href: '/sites/Intranet/SitePages/Blood-donation.aspx' },
            { label: 'Beach Clean-ups', href: '/sites/Intranet/SitePages/Beach-Clean-up.aspx' },
            { label: 'Together We Run', href: '/sites/Intranet/SitePages/Together-we-Run.aspx' },
            { label: 'Events', href: 'https://starbulk.sharepoint.com/sites/Intranet/SitePages/Highlights.aspx' }
          ]
        }
      ]
    ]
  },
  {
    label: 'Workplace',
    columns: [
      [
        {
          items: [
            { label: 'SAP', href: '/sites/Intranet/SitePages/SAP---HRMS.aspx' },
            { label: 'Vacation & Absences', href: '/sites/Intranet/SitePages/Vacation-&-Absences.aspx' },
            { label: 'IT Support', href: '/sites/Intranet/SitePages/IT-Support---Helpdesk.aspx' },
            { label: 'Company Mobile & Laptop', href: '/sites/Intranet/SitePages/Company-Mobile.aspx' },
            { label: 'Employee Discounts', href: '/sites/Intranet/SitePages/Employee-Discounts.aspx' },
            { label: 'Health Insurance', href: '/sites/Intranet/SitePages/Health-Insurance.aspx' }
          ]
        }
      ],
      [
        {
          items: [
            { label: 'Mental Health Hotline', href: '/sites/Intranet/SitePages/Mental-Health-Hotline.aspx' },
            { label: 'Education Allowance', href: '/sites/Intranet/SitePages/Education-Allowance.aspx' },
            { label: 'Brand & Identity Guidelines', href: '/sites/Intranet/SitePages/Corporate-Identity.aspx' },
            { label: 'Buildings & Facilities', href: '/sites/Intranet/SitePages/Buildings-&-Facilities.aspx' },
            { label: 'Building Maintenance', href: '/sites/Intranet/SitePages/Building-Maintenance.aspx' }
          ]
        }
      ]
    ]
  },
  {
    label: 'Culture',
    columns: [
      [
        {
          items: [
            { label: 'Policies', href: '/sites/Intranet/SitePages/Regulations-and-Policies.aspx' },
            { label: 'Data Protection', href: '/sites/Intranet/SitePages/Data-Protection---GDPR.aspx' },
            { label: 'ESG', href: '/sites/Intranet/SitePages/Environmental,-Social-&-Governance.aspx' },
            { label: 'Whistleblowing Platform', href: '/sites/Intranet/SitePages/Whistleblower-Platform.aspx' },
            { label: 'Health & Safety', href: '/sites/Intranet/SitePages/Health,-Safety-and-Emergency.aspx' }
          ]
        }
      ]
    ]
  },
  {
    label: 'Newsroom',
    href: '/sites/Intranet/SitePages/Internal-Company-Announcements.aspx'
  },
  {
    label: 'Fleet',
    href: '/sites/Intranet/SitePages/Fleet.aspx'
  }
];
