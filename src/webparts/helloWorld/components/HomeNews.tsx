import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './HomeNews.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import NewsCarousel from './NewsCarousel';

export interface IHomeNewsProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  /** Site Pages URLs (or file names) pinned for everyone; they stay in place while the other cards rotate. */
  pinnedNews: string[];
}

export default class HomeNews extends React.Component<IHomeNewsProps> {
  public componentDidMount(): void {
    ensureInterFont();
  }

  public render(): React.ReactElement {
    return (
      <div className={styles.page}>
        <div className={styles.main}>
          <NewsCarousel spHttpClient={this.props.spHttpClient} siteUrl={this.props.siteUrl} pinned={this.props.pinnedNews} />
        </div>
      </div>
    );
  }
}
