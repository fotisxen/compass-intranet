import * as React from 'react';
import styles from './WorkAnniversaries.module.scss';
import { anniversaries } from './data/mockData';

const PAGE_SIZE = 4;

export interface IWorkAnniversariesState {
  page: number;
}

export default class WorkAnniversaries extends React.Component<Record<string, never>, IWorkAnniversariesState> {
  constructor(props: Record<string, never>) {
    super(props);
    this.state = { page: 0 };
  }

  private get _pageCount(): number {
    return Math.max(1, Math.ceil(anniversaries.length / PAGE_SIZE));
  }

  private _prev = (): void => {
    this.setState(prev => ({ page: (prev.page - 1 + this._pageCount) % this._pageCount }));
  };

  private _next = (): void => {
    this.setState(prev => ({ page: (prev.page + 1) % this._pageCount }));
  };

  public render(): React.ReactElement {
    const { page } = this.state;
    const visible = anniversaries.slice(page * PAGE_SIZE, page * PAGE_SIZE + PAGE_SIZE);

    return (
      <div className={styles.card}>
        <h3 className={styles.title}>Work Anniversaries</h3>

        {this._pageCount > 1 && (
          <>
            <button className={`${styles.arrowButton} ${styles.arrowLeft}`} onClick={this._prev} aria-label="Previous">←</button>
            <button className={`${styles.arrowButton} ${styles.arrowRight}`} onClick={this._next} aria-label="Next">→</button>
          </>
        )}

        <div className={styles.list}>
          {visible.map(a => (
            <div className={styles.row} key={a.name}>
              <span className={styles.avatar}>{a.initials}</span>
              <div className={styles.name}>
                <div className={styles.nameText}>{a.name}</div>
                <div className={styles.roleText}>{a.role}</div>
              </div>
              <span className={styles.years}>{a.years}Y</span>
            </div>
          ))}
        </div>
      </div>
    );
  }
}
