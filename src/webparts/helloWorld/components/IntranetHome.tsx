import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './IntranetHome.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import PublicHolidays from './PublicHolidays';
import EventsWidget from './EventsWidget';
import FleetMap from './FleetMap';
import NewsSection from './NewsSection';
import PeopleSidebar from './PeopleSidebar';
import PersonSpotlightCard from './PersonSpotlightCard';
import WorkAnniversaries from './WorkAnniversaries';
import { welcomeAboard, promotion } from './data/mockData';

export interface IIntranetHomeProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  fleetApiUrl?: string;
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

          <div className={styles.heroRow}>
            <EventsWidget spHttpClient={this.props.spHttpClient} siteUrl={this.props.siteUrl} />
            <div className={styles.mapColumn}>
              <FleetMap fleetApiUrl={this.props.fleetApiUrl} />
            </div>
          </div>

          <div className={styles.section}>
            <NewsSection spHttpClient={this.props.spHttpClient} siteUrl={this.props.siteUrl} />
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
