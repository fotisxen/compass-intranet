export interface INavGroup {
  heading?: string;
  items: string[];
}

export type INavColumn = INavGroup[];

export interface INavItem {
  label: string;
  columns?: INavColumn[];
}

export const NAV_ITEMS: INavItem[] = [
  {
    label: 'Company Profile',
    columns: [
      [
        {
          items: ['Starbulk Overview', 'Policies', 'Environmental, Social & Governance', 'Buildings & Facilities']
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
          items: ['Welcome Aboard!', 'Health & Safety', 'Open Positions', 'Vacation & Absences', 'Building Maintenance']
        },
        {
          heading: 'VOLUNTEERING',
          items: ['Blood Donation', 'Beach Clean-ups', 'Together We Run']
        }
      ],
      [
        {
          heading: 'EMPLOYEE TOOLS & RESOURCES',
          items: ['SAP', 'Whistleblowing Platform', 'Data Protection', 'Company Mobile & Laptop', 'IT Support', 'Brand & Identity Guidelines']
        },
        {
          heading: 'WELLBEING & BENEFITS',
          items: ['Health Insurance', 'Mental Health Hotline', 'Education Allowance', 'Employee Discounts']
        }
      ]
    ]
  },
  {
    label: 'Newsroom',
    columns: [
      [
        {
          items: ['Corporate', 'Fleet', 'People & Culture', 'Open Positions']
        }
      ]
    ]
  },
  {
    label: 'Fleet'
  }
];
