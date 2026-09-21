import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './PublicHolidays.module.scss';
import { type IHoliday } from './data/holidaysMockData';
import { loadUpcomingHolidays } from './holidaysApi';

export interface IPublicHolidaysProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  /** Shifts the whole widget right (or left, negative) — editable from the web part's property pane since the real page's exact alignment can't be verified until it's live. */
  offsetX?: number;
}

export interface IPublicHolidaysState {
  holidays: IHoliday[];
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
    const holidays = await loadUpcomingHolidays(spHttpClient, siteUrl);
    if (holidays.length > 0) {
      this.setState({ holidays });
    }
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
