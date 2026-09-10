import * as React from 'react';
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

export default class IntranetHome extends React.Component {
  public componentDidMount(): void {
    ensureInterFont();
  }

  public render(): React.ReactElement {
    return (
      <div className={styles.page}>
        <div className={styles.main}>
          <div className={styles.holidaysRow}>
            <PublicHolidays />
          </div>

          <div className={styles.heroRow}>
            <EventsWidget />
            <div className={styles.mapColumn}>
              <FleetMap />
            </div>
          </div>

          <div className={styles.section}>
            <NewsSection />
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
