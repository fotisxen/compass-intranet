import * as React from 'react';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './PublicHolidays.module.scss';
import { type IHoliday } from './data/holidaysMockData';

export interface IPublicHolidaysProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  /** Shifts the whole widget right (or left, negative) — editable from the web part's property pane since the real page's exact alignment can't be verified until it's live. */
  offsetX?: number;
}

export interface IPublicHolidaysState {
  holidays: IHoliday[];
}

// The real "Public Holidays" widget on the live site reads from a plain
// SharePoint Events list (Title + EventDate) — this is that list's GUID,
// found in the page's own web part configuration.
const PUBLIC_HOLIDAYS_LIST_ID = 'd6e5df27-ce8c-4770-9b39-e58c91c50d4b';
// Only the 2 nearest holidays (today or later) — the query below already
// filters to >= today and sorts ascending, so this just caps the count.
const PAGE_SIZE = 2;

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
    this.state = { holidays: [] };
  }

  public componentDidMount(): void {
    this._loadHolidays().catch(() => {
      // Real list couldn't be loaded — leave holidays empty rather than
      // show fabricated ones.
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
    const { offsetX } = this.props;

    return (
      <div className={styles.zone} style={{ marginLeft: `calc(-16px + ${offsetX ?? 0}px)` }}>
        <div className={styles.outer}>
        <div className={styles.widget}>
          <div className={styles.heading}>PUBLIC HOLIDAYS</div>
          <div className={styles.list}>
            {this.state.holidays.length > 0 ? (
              this.state.holidays.map(h => (
                <div className={styles.row} key={h.date + h.label}>
                  <span className={styles.date}>{h.date}</span>
                  <span className={styles.label}>{h.label}</span>
                </div>
              ))
            ) : (
              <div className={styles.emptyRow}>No upcoming holidays</div>
            )}
          </div>
        </div>
        </div>
      </div>
    );
  }
}
