import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldCollectionData, CustomCollectionFieldType, type ICustomCollectionField } from '@pnp/spfx-property-controls';

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import {
  type IPersonSpotlightItem,
  type IAnniversaryItem,
  mapPersonItems,
  mapAnniversaryItems,
  DEFAULT_WELCOME_ABOARD,
  DEFAULT_PROMOTIONS,
  DEFAULT_ANNIVERSARIES
} from './HelloWorldProperties';

export interface IHelloWorldWebPartProps {
  welcomeAboard: IPersonSpotlightItem[];
  promotions: IPersonSpotlightItem[];
  anniversaries: IAnniversaryItem[];
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  protected onInit(): Promise<void> {
    // Backfills an instance placed before these properties existed, so it
    // keeps showing the same content it always did instead of going blank.
    this.properties.welcomeAboard = this.properties.welcomeAboard || DEFAULT_WELCOME_ABOARD;
    this.properties.promotions = this.properties.promotions || DEFAULT_PROMOTIONS;
    this.properties.anniversaries = this.properties.anniversaries || DEFAULT_ANNIVERSARIES;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(HelloWorld, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      welcomeAboard: mapPersonItems(this.properties.welcomeAboard),
      promotions: mapPersonItems(this.properties.promotions),
      anniversaries: mapAnniversaryItems(this.properties.anniversaries)
    });
    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Shared column set for both person-spotlight cards — Promotions just
    // also uses newTitle, and defaults its field labels differently.
    const personFields = (defaults: { field1: string; field2: string; field3: string }): ICustomCollectionField[] => [
      { id: 'name', title: 'Name', type: CustomCollectionFieldType.string, required: true },
      { id: 'title', title: 'Current role', type: CustomCollectionFieldType.string, required: true },
      { id: 'newTitle', title: 'New role (Promotions only)', type: CustomCollectionFieldType.string },
      { id: 'photoUrl', title: 'Photo URL', type: CustomCollectionFieldType.url, placeholder: 'https://.../SiteAssets/photo.jpg' },
      { id: 'field1Label', title: 'Field 1 label', type: CustomCollectionFieldType.string, defaultValue: defaults.field1 },
      { id: 'field1Value', title: 'Field 1 value', type: CustomCollectionFieldType.string },
      { id: 'field2Label', title: 'Field 2 label', type: CustomCollectionFieldType.string, defaultValue: defaults.field2 },
      { id: 'field2Value', title: 'Field 2 value', type: CustomCollectionFieldType.string },
      { id: 'field3Label', title: 'Field 3 label', type: CustomCollectionFieldType.string, defaultValue: defaults.field3 },
      { id: 'field3Value', title: 'Field 3 value', type: CustomCollectionFieldType.string }
    ];

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: 'Welcome Aboard',
              groupFields: [
                PropertyFieldCollectionData('welcomeAboard', {
                  key: 'welcomeAboard',
                  label: 'People to welcome',
                  panelHeader: 'Welcome Aboard — people',
                  manageBtnLabel: 'Manage people',
                  saveAndAddBtnLabel: 'Save and add another',
                  fields: personFields({ field1: 'COMPANY', field2: 'DEPARTMENT', field3: 'REPORTS TO' }),
                  value: this.properties.welcomeAboard
                })
              ]
            },
            {
              groupName: 'Promotions & Internal Transfers',
              groupFields: [
                PropertyFieldCollectionData('promotions', {
                  key: 'promotions',
                  label: 'People with a move',
                  panelHeader: 'Promotions & Internal Transfers — people',
                  manageBtnLabel: 'Manage people',
                  saveAndAddBtnLabel: 'Save and add another',
                  fields: personFields({ field1: 'MOVE TYPE', field2: 'DEPARTMENT', field3: 'REPORTS TO' }),
                  value: this.properties.promotions
                })
              ]
            },
            {
              groupName: 'Work Anniversaries',
              groupFields: [
                PropertyFieldCollectionData('anniversaries', {
                  key: 'anniversaries',
                  label: 'Anniversaries (shown 4 per page)',
                  panelHeader: 'Work Anniversaries — people',
                  manageBtnLabel: 'Manage people',
                  saveAndAddBtnLabel: 'Save and add another',
                  fields: [
                    { id: 'name', title: 'Name', type: CustomCollectionFieldType.string, required: true },
                    { id: 'role', title: 'Role', type: CustomCollectionFieldType.string, required: true },
                    { id: 'years', title: 'Years', type: CustomCollectionFieldType.number, required: true }
                  ],
                  value: this.properties.anniversaries
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
