import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './IntranetHome.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import NewsroomSidebar from '../../../shared/components/NewsroomSidebar';
import NewsCard from '../../../shared/components/NewsCard';
import PeopleSidebar from './PeopleSidebar';
import PersonSpotlightCard from './PersonSpotlightCard';
import WorkAnniversaries from './WorkAnniversaries';
import type { IPersonSpotlight, IAnniversary } from './data/mockData';

export interface IIntranetHomeProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  welcomeAboard: IPersonSpotlight[];
  promotions: IPersonSpotlight[];
  anniversaries: IAnniversary[];
}

export interface IIntranetHomeState {
  selectedNewsCategory?: string;
}

export default class IntranetHome extends React.Component<IIntranetHomeProps, IIntranetHomeState> {
  constructor(props: IIntranetHomeProps) {
    super(props);
    this.state = {};
  }

  public componentDidMount(): void {
    ensureInterFont();
  }

  private _onSelectNewsCategory = (label: string): void => {
    this.setState({ selectedNewsCategory: label });
  };

  public render(): React.ReactElement {
    const { selectedNewsCategory } = this.state;

    return (
      <div className={styles.page}>
        <div className={styles.main}>
          <div className={styles.section}>
            <div className={styles.newsWrap}>
              <NewsroomSidebar selectedLabel={selectedNewsCategory} onSelect={this._onSelectNewsCategory} />
              <div className={styles.newsGrid}>
                <NewsCard spHttpClient={this.props.spHttpClient} siteUrl={this.props.siteUrl} position={1} categoryFilter={selectedNewsCategory} />
                <NewsCard spHttpClient={this.props.spHttpClient} siteUrl={this.props.siteUrl} position={2} categoryFilter={selectedNewsCategory} />
                <NewsCard spHttpClient={this.props.spHttpClient} siteUrl={this.props.siteUrl} position={3} categoryFilter={selectedNewsCategory} />
              </div>
            </div>
          </div>

          <div className={styles.peopleRow}>
            <PeopleSidebar />
            <div className={styles.peopleContent}>
              <div className={styles.peopleColumn}>
                <PersonSpotlightCard heading="Welcome Aboard" headingFontSize={20} background="#F4F8FB" people={this.props.welcomeAboard} />
                <PersonSpotlightCard heading="Promotions & Internal Transfers" headingFontSize={18} background="#C7D7E6" people={this.props.promotions} />
              </div>
              <div className={styles.anniversariesColumn}>
                <WorkAnniversaries people={this.props.anniversaries} />
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
