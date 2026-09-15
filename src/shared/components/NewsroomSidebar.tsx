import * as React from 'react';
import styles from './NewsroomSidebar.module.scss';
import { newsCategories } from './data/newsMockData';

export interface INewsroomSidebarState {
  selectedLabel?: string;
}

export default class NewsroomSidebar extends React.Component<Record<string, never>, INewsroomSidebarState> {
  constructor(props: Record<string, never>) {
    super(props);
    this.state = {};
  }

  private _select = (label: string): ((e: React.MouseEvent) => void) => (e: React.MouseEvent): void => {
    e.preventDefault();
    this.setState({ selectedLabel: label });
  };

  public render(): React.ReactElement {
    const { selectedLabel } = this.state;

    return (
      <div className={styles.sidebar}>
        <h3 className={styles.sidebarTitle}>Newsroom</h3>
        <div className={styles.tabs}>
          {newsCategories.map(cat => {
            const isSelected = cat.label === selectedLabel;
            return (
              <div key={cat.label} className={styles.tabRow}>
                <a
                  href="#"
                  onClick={this._select(cat.label)}
                  className={[
                    cat.highlighted ? styles.tabHighlighted : styles.tab,
                    isSelected ? styles.tabActive : ''
                  ].join(' ')}
                >
                  {cat.label}
                </a>
                {isSelected && (
                  <a
                    href={cat.href || '#'}
                    className={styles.arrowButton}
                    aria-label={`Go to ${cat.label}`}
                  >
                    →
                  </a>
                )}
              </div>
            );
          })}
        </div>
      </div>
    );
  }
}
