import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
// Deep import (not the package barrel) — the barrel re-exports every control
// in the library, which bloats this bundle by ~9MB for just one control.
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'LogoGridFeatureWebPartStrings';
import LogoGridFeature, { type ILogoItem } from './components/LogoGridFeature';

export interface ILogoGridFeatureWebPartProps {
  imageUrl?: string;
  logos: ILogoItem[];
  title: string;
  bodyText: string;
}

export default class LogoGridFeatureWebPart extends BaseClientSideWebPart<ILogoGridFeatureWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.title = this.properties.title ?? 'Title';
    this.properties.bodyText = this.properties.bodyText ?? 'Text';
    this.properties.logos = this.properties.logos ?? [];
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(LogoGridFeature, { ...this.properties });
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
    return {
      pages: [{
        header: { description: strings.PropertyPaneDescription },
        groups: [{
          groupName: strings.BasicGroupName,
          groupFields: [
            PropertyPaneTextField('imageUrl', { label: strings.ImageUrlFieldLabel, placeholder: 'https://.../SiteAssets/photo.jpg' }),
            PropertyPaneTextField('title', { label: strings.TitleFieldLabel }),
            PropertyPaneTextField('bodyText', { label: strings.BodyTextFieldLabel, multiline: true }),
            PropertyFieldCollectionData('logos', {
              key: 'logos',
              label: strings.LogosFieldLabel,
              panelHeader: 'Logos',
              manageBtnLabel: 'Manage logos',
              saveAndAddBtnLabel: 'Save and add another',
              fields: [
                { id: 'logoUrl', title: 'Logo image URL', type: CustomCollectionFieldType.url, required: true },
                { id: 'linkUrl', title: 'Link URL (optional)', type: CustomCollectionFieldType.url }
              ],
              value: this.properties.logos
            })
          ]
        }]
      }]
    };
  }
}
