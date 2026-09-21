import * as React from 'react';
import styles from './OpenPositions.module.scss';

const PAGE_SIZE = 4;
const JOB_LISTINGS_API_URL = 'https://bpcstarbulkwebapi.azurewebsites.net/api/GetJobListings';

interface IJobListing {
  title: string;
  department: string;
}

interface IJobListingApiEntry {
  title?: string;
  department?: string;
}

export interface IOpenPositionsState {
  jobs: IJobListing[];
  page: number;
}

export default class OpenPositions extends React.Component<Record<string, never>, IOpenPositionsState> {
  constructor(props: Record<string, never>) {
    super(props);
    this.state = { jobs: [], page: 0 };
  }

  public componentDidMount(): void {
    this._loadJobs().catch(() => {
      // Listings unavailable — leave the card empty rather than show
      // made-up positions.
    });
  }

  private async _loadJobs(): Promise<void> {
    const response = await fetch(JOB_LISTINGS_API_URL);
    if (!response.ok) {
      return;
    }

    const data: IJobListingApiEntry[] = await response.json();
    if (!Array.isArray(data)) {
      return;
    }

    const jobs: IJobListing[] = data
      .filter(e => e.title && e.title.trim())
      .map(e => ({ title: (e.title || '').trim(), department: (e.department || '').trim() }));

    this.setState({ jobs, page: 0 });
  }

  private get _pageCount(): number {
    return Math.max(1, Math.ceil(this.state.jobs.length / PAGE_SIZE));
  }

  private _prev = (): void => {
    this.setState(prev => ({ page: (prev.page - 1 + this._pageCount) % this._pageCount }));
  };

  private _next = (): void => {
    this.setState(prev => ({ page: (prev.page + 1) % this._pageCount }));
  };

  public render(): React.ReactElement {
    const { jobs, page } = this.state;
    const visible = jobs.slice(page * PAGE_SIZE, page * PAGE_SIZE + PAGE_SIZE);

    return (
      <div className={styles.card}>
        <h3 className={styles.title}>Open Positions</h3>

        {this._pageCount > 1 && (
          <>
            <button className={`${styles.arrowButton} ${styles.arrowLeft}`} onClick={this._prev} aria-label="Previous">←</button>
            <button className={`${styles.arrowButton} ${styles.arrowRight}`} onClick={this._next} aria-label="Next">→</button>
          </>
        )}

        <div className={styles.list}>
          {visible.map((j, i) => (
            <div className={styles.row} key={`${j.title}-${page}-${i}`}>
              <div className={styles.titleText}>{j.title}</div>
              {j.department && <div className={styles.departmentText}>{j.department}</div>}
            </div>
          ))}
        </div>
      </div>
    );
  }
}
