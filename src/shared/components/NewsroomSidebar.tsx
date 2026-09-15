import * as React from 'react';
import styles from './NewsroomSidebar.module.scss';
import { newsCategories } from './data/newsMockData';

export interface INewsroomSidebarProps {
  /**
   * Controlled selection — pass both to filter an adjacent News Card list
   * (e.g. composed inside IntranetHome). Omit both for standalone web part
   * use: category labels then behave like the real production Newsroom —
   * plain links that open the matching announcements page in a new tab,
   * no in-page filtering.
   */
  selectedLabel?: string;
  onSelect?: (label: string) => void;
}

export default class NewsroomSidebar extends React.Component<INewsroomSidebarProps> {
  private _select = (label: string): ((e: React.MouseEvent) => void) => (e: React.MouseEvent): void => {
    e.preventDefault();
    this.props.onSelect!(label);
  };

  public render(): React.ReactElement {
    const { selectedLabel, onSelect } = this.props;
    const isFiltering = !!onSelect;

    return (
      <div className={styles.sidebar}>
        <h3 className={styles.sidebarTitle}>Newsroom</h3>
        <div className={styles.tabs}>
          {newsCategories.map(cat => {
            const isSelected = isFiltering && cat.label === selectedLabel;
            return (
              <div key={cat.label} className={styles.tabRow}>
                <a
                  href={isFiltering ? '#' : cat.href || '#'}
                  target={isFiltering ? undefined : '_blank'}
                  rel={isFiltering ? undefined : 'noopener noreferrer'}
                  onClick={isFiltering ? this._select(cat.label) : undefined}
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
                    target="_blank"
                    rel="noopener noreferrer"
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
