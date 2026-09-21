import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './IntranetHome.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import NewsCard from '../../../shared/components/NewsCard';
import PersonSpotlightCard from './PersonSpotlightCard';
import WorkAnniversaries from './WorkAnniversaries';
import OpenPositions from './OpenPositions';
import type { IPersonSpotlight, IAnniversary } from './data/mockData';

export interface IIntranetHomeProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  welcomeAboard: IPersonSpotlight[];
  promotions: IPersonSpotlight[];
  anniversaries: IAnniversary[];
}

export default class IntranetHome extends React.Component<IIntranetHomeProps> {
  public componentDidMount(): void {
    ensureInterFont();
  }

  public render(): React.ReactElement {
    return (
      <div className={styles.page}>
        <div className={styles.main}>
          <div className={styles.section}>
            <div className={styles.newsGrid}>
              <NewsCard spHttpClient={this.props.spHttpClient} siteUrl={this.props.siteUrl} position={1} />
              <NewsCard spHttpClient={this.props.spHttpClient} siteUrl={this.props.siteUrl} position={2} />
              <NewsCard spHttpClient={this.props.spHttpClient} siteUrl={this.props.siteUrl} position={3} />
            </div>
          </div>

          <div className={styles.peopleContent}>
            <div className={styles.peopleColumn}>
              <PersonSpotlightCard heading="Welcome Aboard" headingFontSize={20} background="#F4F8FB" people={this.props.welcomeAboard} />
              <PersonSpotlightCard heading="Promotions & Internal Transfers" headingFontSize={20} background="#C7D7E6" people={this.props.promotions} />
            </div>
            <div className={styles.openPositionsColumn}>
              <OpenPositions />
            </div>
            <div className={styles.anniversariesColumn}>
              <WorkAnniversaries people={this.props.anniversaries} />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
