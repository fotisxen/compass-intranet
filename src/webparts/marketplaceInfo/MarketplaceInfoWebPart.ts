import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneSlider, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import MarketplaceInfo from './components/MarketplaceInfo';
import { DEFAULT_EXPIRY_DAYS } from '../../shared/marketplace/marketplaceApi';

export interface IMarketplaceInfoWebPartProps {
  rulesTitle?: string;
  rulesText?: string;
  contactLabel?: string;
  contactEmail?: string;
  contactSubject?: string;
  contactBody?: string;
  manageUrl?: string;
  expiryDays?: number;
}

const DEFAULT_RULES = [
  'Personal items only, no running a business',
  'Nothing prohibited or unsafe',
  'Listings expire after 30 days',
  'Deals are between colleagues; the company is not responsible for transactions'
].join('\n');

export default class MarketplaceInfoWebPart extends BaseClientSideWebPart<IMarketplaceInfoWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.rulesTitle = this.properties.rulesTitle ?? 'MARKETPLACE RULES';
    this.properties.rulesText = this.properties.rulesText ?? DEFAULT_RULES;
    this.properties.contactLabel = this.properties.contactLabel ?? 'Questions? Contact HR';
    this.properties.contactSubject = this.properties.contactSubject ?? 'Marketplace question';
    this.properties.contactBody = this.properties.contactBody ?? 'Hello HR,\n\nI have a question about the Marketplace:\n\n';
    this.properties.expiryDays = this.properties.expiryDays ?? DEFAULT_EXPIRY_DAYS;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(MarketplaceInfo, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      isEditMode: this.displayMode === DisplayMode.Edit,
      rulesTitle: this.properties.rulesTitle || '',
      rulesText: this.properties.rulesText || '',
      contactLabel: this.properties.contactLabel || '',
      contactEmail: (this.properties.contactEmail || '').trim(),
      contactSubject: this.properties.contactSubject || '',
      contactBody: this.properties.contactBody || '',
      manageUrl: (this.properties.manageUrl || '').trim(),
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
        header: { description: 'Rules text: one rule per line, each shown as a bullet.' },
        groups: [
          {
            groupName: 'Marketplace rules',
            groupFields: [
              PropertyPaneTextField('rulesTitle', { label: 'Heading' }),
              PropertyPaneTextField('rulesText', { label: 'Rules (one per line)', multiline: true, rows: 8 })
            ]
          },
          {
            groupName: 'Contact HR',
            groupFields: [
              PropertyPaneTextField('contactLabel', { label: 'Link text' }),
              PropertyPaneTextField('contactEmail', { label: 'HR email address', description: 'The link stays hidden until this is filled in.' }),
              PropertyPaneTextField('contactSubject', { label: 'Pre-filled subject' }),
              PropertyPaneTextField('contactBody', { label: 'Pre-filled message', multiline: true, rows: 5 })
            ]
          },
          {
            groupName: 'My listings',
            groupFields: [
              PropertyPaneTextField('manageUrl', { label: 'Manage listings link (optional)', description: 'Leave empty to open the Compass Marketplace list.' }),
              PropertyPaneSlider('expiryDays', { label: 'Listings expire after (days, 0 = never) — same as the Listings web part', min: 0, max: 180, step: 5, value: this.properties.expiryDays ?? DEFAULT_EXPIRY_DAYS })
            ]
          }
        ]
      }]
    };
  }
}
