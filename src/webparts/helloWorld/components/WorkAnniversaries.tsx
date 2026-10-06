import * as React from 'react';
import styles from './WorkAnniversaries.module.scss';
import type { IAnniversary } from './data/mockData';

const PAGE_SIZE = 4;

export interface IWorkAnniversariesProps {
  people: IAnniversary[];
}

export interface IWorkAnniversariesState {
  page: number;
}

export default class WorkAnniversaries extends React.Component<IWorkAnniversariesProps, IWorkAnniversariesState> {
  constructor(props: IWorkAnniversariesProps) {
    super(props);
    this.state = { page: 0 };
  }

  private get _pageCount(): number {
    return Math.max(1, Math.ceil(this.props.people.length / PAGE_SIZE));
  }

  private _prev = (): void => {
    this.setState(prev => ({ page: (prev.page - 1 + this._pageCount) % this._pageCount }));
  };

  private _next = (): void => {
    this.setState(prev => ({ page: (prev.page + 1) % this._pageCount }));
  };

  public render(): React.ReactElement {
    const { page } = this.state;
    const visible = this.props.people.slice(page * PAGE_SIZE, page * PAGE_SIZE + PAGE_SIZE);

    return (
      <div className={styles.card}>
        <h3 className={styles.title}>Work Anniversaries</h3>

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
          {/* Invisible filler rows keep the list as tall as a full page, so the arrows below stay put on the last (shorter) page. */}
          {this._pageCount > 1 && Array.from({ length: PAGE_SIZE - visible.length }, (_, i) => (
            <div className={`${styles.row} ${styles.rowPlaceholder}`} key={`filler-${i}`} aria-hidden="true">
              <span className={styles.avatar} />
              <div className={styles.name}>
                <div className={styles.nameText}>&nbsp;</div>
                <div className={styles.roleText}>&nbsp;</div>
              </div>
            </div>
          ))}
        </div>

        {this._pageCount > 1 && (
          <div className={styles.arrowsRow}>
            <button className={styles.arrowButton} onClick={this._prev} aria-label="Previous">←</button>
            <button className={styles.arrowButton} onClick={this._next} aria-label="Next">→</button>
          </div>
        )}
      </div>
    );
  }
}
