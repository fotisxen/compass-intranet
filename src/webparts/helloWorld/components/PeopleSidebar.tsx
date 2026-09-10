import * as React from 'react';
import styles from './PeopleSidebar.module.scss';

export default class PeopleSidebar extends React.Component {
  public render(): React.ReactElement {
    return (
      <div className={styles.sidebar}>
        <h3 className={styles.title}>Our People</h3>
      </div>
    );
  }
}
