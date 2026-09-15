import * as React from 'react';
import styles from './NewsroomSidebar.module.scss';
import { newsCategories } from './data/newsMockData';

// Matches the real production Newsroom: category labels are plain links
// that open the matching announcements page in a new tab. No in-page
// filtering — the promoted "Site Pages" items have no category field to
// filter by yet (see NewsCard for the fetch this would need to match).
export default class NewsroomSidebar extends React.Component {
  public render(): React.ReactElement {
    return (
      <div className={styles.sidebar}>
        <h3 className={styles.sidebarTitle}>Newsroom</h3>
        <div className={styles.tabs}>
          {newsCategories.map(cat => (
            <a
              key={cat.label}
              href={cat.href || '#'}
              target="_blank"
              rel="noopener noreferrer"
              className={cat.highlighted ? styles.tabHighlighted : styles.tab}
            >
              {cat.label}
            </a>
          ))}
        </div>
      </div>
    );
  }
}
