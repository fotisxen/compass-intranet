import * as React from 'react';
import styles from './HomePeople.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import PersonSpotlightCard from './PersonSpotlightCard';
import WorkAnniversaries from './WorkAnniversaries';
import OpenPositions from './OpenPositions';
import type { IPersonSpotlight, IAnniversary } from './data/mockData';
import { welcomeAboard as defaultWelcomeAboard, promotion as defaultPromotions, anniversaries as defaultAnniversaries } from './data/mockData';

export interface IHomePeopleProps {
  /** Omit for the standalone preview — falls back to the original hand-authored sample content. */
  welcomeAboard?: IPersonSpotlight[];
  promotions?: IPersonSpotlight[];
  anniversaries?: IAnniversary[];
}

export default class HomePeople extends React.Component<IHomePeopleProps> {
  public componentDidMount(): void {
    ensureInterFont();
  }

  public render(): React.ReactElement {
    return (
      <div className={styles.page}>
        <div className={styles.main}>
          <div className={styles.peopleContent}>
            <div className={styles.peopleColumn}>
              <PersonSpotlightCard heading="Welcome Aboard" headingFontSize={22} background="#F4F8FB" people={this.props.welcomeAboard || defaultWelcomeAboard} />
              <PersonSpotlightCard heading="Promotions & Internal Transfers" headingFontSize={22} background="#C7D7E6" people={this.props.promotions || defaultPromotions} />
            </div>
            <div className={styles.openPositionsColumn}>
              <OpenPositions />
            </div>
            <div className={styles.anniversariesColumn}>
              <WorkAnniversaries people={this.props.anniversaries || defaultAnniversaries} />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
