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
    label: 'Company Profile',
    columns: [
      [
        {
          items: [
            { label: 'Starbulk Overview', href: '/sites/Intranet/SitePages/Corporate-Profile.aspx' },
            { label: 'Policies', href: '/sites/Intranet/SitePages/Regulations-and-Policies.aspx' },
            { label: 'Environmental, Social & Governance', href: '/sites/Intranet/SitePages/Environmental,-Social-&-Governance.aspx' },
            { label: 'Buildings & Facilities', href: '/sites/Intranet/SitePages/Buildings-&-Facilities.aspx' }
          ]
        }
      ]
    ]
  },
  {
    label: 'My Workplace',
    columns: [
      [
        {
          heading: 'HUMAN RESOURCES',
          headingHref: '/sites/Intranet/SitePages/Human-Resources.aspx',
          items: [
            { label: 'Welcome Aboard!', href: '/sites/Intranet/SitePages/Onboarding---Welcome-Aboard!.aspx' },
            { label: 'Health & Safety', href: '/sites/Intranet/SitePages/Health,-Safety-and-Emergency.aspx' },
            { label: 'Open Positions', href: '/sites/Intranet/SitePages/Career-Opportunities.aspx' },
            { label: 'Vacation & Absences', href: '/sites/Intranet/SitePages/Vacation-&-Absences.aspx' },
            { label: 'Building Maintenance', href: '/sites/Intranet/SitePages/Building-Maintenance.aspx' }
          ]
        },
        {
          heading: 'VOLUNTEERING',
          headingHref: '/sites/Intranet/SitePages/Volunteering.aspx',
          items: [
            { label: 'Blood Donation', href: '/sites/Intranet/SitePages/Blood-donation.aspx' },
            { label: 'Beach Clean-ups', href: '/sites/Intranet/SitePages/Beach-Clean-up.aspx' },
            { label: 'Together We Run', href: '/sites/Intranet/SitePages/Together-we-Run.aspx' }
          ]
        }
      ],
      [
        {
          heading: 'EMPLOYEE TOOLS & RESOURCES',
          headingHref: '/sites/Intranet/SitePages/Digital-Hub.aspx',
          items: [
            { label: 'SAP', href: '/sites/Intranet/SitePages/SAP---HRMS.aspx' },
            { label: 'Whistleblowing Platform', href: '/sites/Intranet/SitePages/Whistleblower-Platform.aspx' },
            { label: 'Data Protection', href: '/sites/Intranet/SitePages/Data-Protection---GDPR.aspx' },
            { label: 'Company Mobile & Laptop', href: '/sites/Intranet/SitePages/Company-Mobile.aspx' },
            { label: 'IT Support', href: '/sites/Intranet/SitePages/IT-Support---Helpdesk.aspx' },
            { label: 'Brand & Identity Guidelines', href: '/sites/Intranet/SitePages/Corporate-Identity.aspx' }
          ]
        },
        {
          heading: 'WELLBEING & BENEFITS',
          headingHref: '/sites/Intranet/SitePages/Wellbeing-&-Benefits.aspx',
          items: [
            { label: 'Health Insurance', href: '/sites/Intranet/SitePages/Health-Insurance.aspx' },
            { label: 'Mental Health Hotline', href: '/sites/Intranet/SitePages/Mental-Health-Hotline.aspx' },
            { label: 'Education Allowance', href: '/sites/Intranet/SitePages/Education-Allowance.aspx' },
            { label: 'Employee Discounts', href: '/sites/Intranet/SitePages/Employee-Discounts.aspx' }
          ]
        }
      ]
    ]
  },
  {
    label: 'Newsroom',
    columns: [
      [
        {
          items: [
            { label: 'Internal Company Announcements', href: '/sites/Intranet/SitePages/Internal-Company-Announcements.aspx' },
            { label: 'Organizational Developments', href: '/sites/Intranet/SitePages/Organizational-Developments.aspx' }
          ]
        }
      ]
    ]
  },
  {
    label: 'Fleet',
    href: '/sites/Intranet/SitePages/Fleet.aspx'
  }
];
