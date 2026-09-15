import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './IntranetHome.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import PublicHolidays from './PublicHolidays';
import NewsroomSidebar from '../../../shared/components/NewsroomSidebar';
import NewsCard from '../../../shared/components/NewsCard';
import PeopleSidebar from './PeopleSidebar';
import PersonSpotlightCard from './PersonSpotlightCard';
import WorkAnniversaries from './WorkAnniversaries';
import { welcomeAboard, promotion } from './data/mockData';

export interface IIntranetHomeProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
}

export default class IntranetHome extends React.Component<IIntranetHomeProps> {
  public componentDidMount(): void {
    ensureInterFont();
  }

  public render(): React.ReactElement {
    return (
      <div className={styles.page}>
        <div className={styles.main}>
          <div className={styles.holidaysRow}>
            <PublicHolidays spHttpClient={this.props.spHttpClient} siteUrl={this.props.siteUrl} />
          </div>

          <div className={styles.section}>
            <div className={styles.newsWrap}>
              <NewsroomSidebar />
              <div className={styles.newsGrid}>
                <NewsCard spHttpClient={this.props.spHttpClient} siteUrl={this.props.siteUrl} position={1} />
                <NewsCard spHttpClient={this.props.spHttpClient} siteUrl={this.props.siteUrl} position={2} />
                <NewsCard spHttpClient={this.props.spHttpClient} siteUrl={this.props.siteUrl} position={3} />
              </div>
            </div>
          </div>

          <div className={styles.peopleRow}>
            <PeopleSidebar />
            <div className={styles.peopleContent}>
              <div className={styles.peopleColumn}>
                <PersonSpotlightCard heading="Welcome Aboard" background="#F4F8FB" person={welcomeAboard} />
                <PersonSpotlightCard heading="Promotions & Internal Transfers" background="#C7D7E6" person={promotion} />
              </div>
              <div className={styles.anniversariesColumn}>
                <WorkAnniversaries />
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
