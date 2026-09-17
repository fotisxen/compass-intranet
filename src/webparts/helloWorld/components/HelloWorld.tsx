import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import IntranetHome from './IntranetHome';
import type { IPersonSpotlight, IAnniversary } from './data/mockData';
import { welcomeAboard as defaultWelcomeAboard, promotion as defaultPromotions, anniversaries as defaultAnniversaries } from './data/mockData';

export interface IHelloWorldProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  /** Omit for the standalone preview — falls back to the original hand-authored sample content. */
  welcomeAboard?: IPersonSpotlight[];
  promotions?: IPersonSpotlight[];
  anniversaries?: IAnniversary[];
}

export default class HelloWorld extends React.Component<IHelloWorldProps> {
  public render(): React.ReactElement {
    return (
      <IntranetHome
        spHttpClient={this.props.spHttpClient}
        siteUrl={this.props.siteUrl}
        welcomeAboard={this.props.welcomeAboard || defaultWelcomeAboard}
        promotions={this.props.promotions || defaultPromotions}
        anniversaries={this.props.anniversaries || defaultAnniversaries}
      />
    );
  }
}
