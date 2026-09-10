import * as React from 'react';
import styles from './NewsSection.module.scss';
import { news, newsCategories } from './data/mockData';

export default class NewsSection extends React.Component {
  public render(): React.ReactElement {
    return (
      <div className={styles.wrap}>
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

        <div className={styles.grid}>
          {news.map(n => (
            <div className={styles.card} key={n.title}>
              <div className={styles.thumb} />
              <div className={styles.body}>
                <h4 className={styles.title}>{n.title}</h4>
                <div className={styles.footerRow}>
                  <span className={styles.date}>{n.date}</span>
                  <button className={styles.pill}>Read more →</button>
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }
}
