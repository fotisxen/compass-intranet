import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import IntranetHome from './IntranetHome';

export interface IHelloWorldProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  fleetApiUrl?: string;
}

export default class HelloWorld extends React.Component<IHelloWorldProps> {
  public render(): React.ReactElement {
    return (
      <IntranetHome
        spHttpClient={this.props.spHttpClient}
        siteUrl={this.props.siteUrl}
        fleetApiUrl={this.props.fleetApiUrl}
      />
    );
  }
}
