import * as React from 'react';
import styles from './PublicHolidays.module.scss';
import { holidays } from './data/mockData';

export default class PublicHolidays extends React.Component {
  public render(): React.ReactElement {
    return (
      <div className={styles.widget}>
        <div className={styles.heading}>PUBLIC HOLIDAYS</div>
        {holidays.map(h => (
          <div className={styles.row} key={h.date}>
            <span className={styles.date}>{h.date}</span>
            <span className={styles.label}>{h.label}</span>
          </div>
        ))}
      </div>
    );
  }
}
