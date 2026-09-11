import * as React from 'react';
import styles from './NewsroomSidebar.module.scss';
import { newsCategories } from './data/newsMockData';

export default class NewsroomSidebar extends React.Component {
  public render(): React.ReactElement {
    return (
      <div className={styles.sidebar}>
        <h3 className={styles.sidebarTitle}>Newsroom</h3>
        <div className={styles.tabs}>
          {newsCategories.map(cat => (
            <a
              key={cat.label}
              href="#"
              onClick={e => e.preventDefault()}
              className={
                cat.highlighted
                  ? styles.tabHighlighted
                  : `${styles.tab} ${cat.label === 'Corporate' ? styles.tabActive : ''}`
              }
            >
              {cat.label}
            </a>
          ))}
        </div>
      </div>
    );
  }
}
