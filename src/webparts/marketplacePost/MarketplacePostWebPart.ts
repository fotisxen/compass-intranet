import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { SPPermission } from '@microsoft/sp-page-context';
import { type IPropertyPaneConfiguration, PropertyPaneSlider, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import MarketplacePost from './components/MarketplacePost';
import { DEFAULT_MAX_PHOTOS } from '../../shared/marketplace/marketplaceApi';

export interface IMarketplacePostWebPartProps {
  eyebrow?: string;
  heading?: string;
  buttonLabel?: string;
  maxPhotos?: number;
}

export default class MarketplacePostWebPart extends BaseClientSideWebPart<IMarketplacePostWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.eyebrow = this.properties.eyebrow ?? 'SELL · GIVE · SWAP';
    this.properties.heading = this.properties.heading ?? "Got something you don't need anymore? Post it in under a minute.";
    this.properties.buttonLabel = this.properties.buttonLabel ?? 'Post an item';
    this.properties.maxPhotos = this.properties.maxPhotos ?? DEFAULT_MAX_PHOTOS;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(MarketplacePost, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      canManageLists: this.context.pageContext.web.permissions.hasPermission(SPPermission.manageLists),
      eyebrow: this.properties.eyebrow || '',
      heading: this.properties.heading || '',
      buttonLabel: this.properties.buttonLabel || 'Post an item',
      maxPhotos: this.properties.maxPhotos || DEFAULT_MAX_PHOTOS
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
        header: { description: 'The card text and the most photos one listing can have.' },
        groups: [{
          groupName: 'Post an item',
          groupFields: [
            PropertyPaneTextField('eyebrow', { label: 'Small heading' }),
            PropertyPaneTextField('heading', { label: 'Text', multiline: true }),
            PropertyPaneTextField('buttonLabel', { label: 'Button label' }),
            PropertyPaneSlider('maxPhotos', { label: 'Most photos per listing', min: 1, max: 8, step: 1, value: this.properties.maxPhotos ?? DEFAULT_MAX_PHOTOS })
          ]
        }]
      }]
    };
  }
}
