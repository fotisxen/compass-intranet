import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { SPPermission } from '@microsoft/sp-page-context';
import { type IPropertyPaneConfiguration, PropertyPaneSlider } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import MarketplaceListings from './components/MarketplaceListings';
import { DEFAULT_EXPIRY_DAYS } from '../../shared/marketplace/marketplaceApi';

export interface IMarketplaceListingsWebPartProps {
  expiryDays?: number;
}

export default class MarketplaceListingsWebPart extends BaseClientSideWebPart<IMarketplaceListingsWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.expiryDays = this.properties.expiryDays ?? DEFAULT_EXPIRY_DAYS;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(MarketplaceListings, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      canManageLists: this.context.pageContext.web.permissions.hasPermission(SPPermission.manageLists),
      expiryDays: this.properties.expiryDays ?? DEFAULT_EXPIRY_DAYS
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'Only approved listings are shown, newest first.' },
        groups: [{
          groupName: 'Listings',
          groupFields: [
            PropertyPaneSlider('expiryDays', {
              label: 'Hide listings after this many days (0 = never)',
              min: 0,
              max: 180,
              step: 5,
              value: this.properties.expiryDays ?? DEFAULT_EXPIRY_DAYS
            })
          ]
        }]
      }]
    };
  }
}
