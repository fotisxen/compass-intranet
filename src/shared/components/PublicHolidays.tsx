import * as React from 'react';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './PublicHolidays.module.scss';
import { holidays as mockHolidays, type IHoliday } from './data/holidaysMockData';

export interface IPublicHolidaysProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
}

export interface IPublicHolidaysState {
  holidays: IHoliday[];
}

// The real "Public Holidays" widget on the live site reads from a plain
// SharePoint Events list (Title + EventDate) — this is that list's GUID,
// found in the page's own web part configuration.
const PUBLIC_HOLIDAYS_LIST_ID = 'd6e5df27-ce8c-4770-9b39-e58c91c50d4b';
const PAGE_SIZE = 5;

interface ISpEventItem {
  Title: string;
  EventDate: string;
}

interface ISpListItemsResponse {
  value: ISpEventItem[];
}

function formatShortDate(iso: string): string {
  return new Date(iso).toLocaleDateString('en-US', { month: 'short', day: 'numeric' }).toUpperCase();
}

export default class PublicHolidays extends React.Component<IPublicHolidaysProps, IPublicHolidaysState> {
  constructor(props: IPublicHolidaysProps) {
    super(props);
    this.state = { holidays: mockHolidays };
  }

  public componentDidMount(): void {
    this._loadHolidays().catch(() => {
      // Real list couldn't be loaded — the mock data already in state
      // stays as a fallback so the widget never renders empty.
    });
  }

  private async _loadHolidays(): Promise<void> {
    const { spHttpClient, siteUrl } = this.props;
    const today = new Date().toISOString();
    const endpoint =
      `${siteUrl}/_api/web/lists(guid'${PUBLIC_HOLIDAYS_LIST_ID}')/items` +
      `?$select=Title,EventDate&$filter=EventDate ge datetime'${today}'&$orderby=EventDate asc&$top=${PAGE_SIZE}`;

    const response: SPHttpClientResponse = await spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    if (!response.ok) {
      return;
    }

    const data: ISpListItemsResponse = await response.json();
    if (!data.value || data.value.length === 0) {
      return;
    }

    const holidays: IHoliday[] = data.value.map(item => ({
      date: formatShortDate(item.EventDate),
      label: item.Title
    }));

    this.setState({ holidays });
  }

  public render(): React.ReactElement {
    return (
      <div className={styles.widget}>
        <div className={styles.heading}>PUBLIC HOLIDAYS</div>
        {this.state.holidays.map(h => (
          <div className={styles.row} key={h.date + h.label}>
            <span className={styles.date}>{h.date}</span>
            <span className={styles.label}>{h.label}</span>
          </div>
        ))}
      </div>
    );
  }
}
