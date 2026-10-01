import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WideNewsCardWebPartStrings';
import WideNewsCard from './components/WideNewsCard';

export interface IWideNewsCardWebPartProps {
  title: string;
  imageUrl?: string;
  dateText: string;
  linkUrl: string;
}

export default class WideNewsCardWebPart extends BaseClientSideWebPart<IWideNewsCardWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.title = this.properties.title ?? 'Fleet renewal programme reaches a new milestone';
    this.properties.dateText = this.properties.dateText ?? 'September 20, 2026';
    this.properties.linkUrl = this.properties.linkUrl ?? '#';
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(WideNewsCard, { ...this.properties });
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
            PropertyPaneTextField('title', { label: strings.TitleFieldLabel, multiline: true }),
            PropertyPaneTextField('imageUrl', { label: strings.ImageUrlFieldLabel, placeholder: 'https://.../SiteAssets/photo.jpg (optional)' }),
            PropertyPaneTextField('dateText', { label: strings.DateTextFieldLabel }),
            PropertyPaneTextField('linkUrl', { label: strings.LinkUrlFieldLabel, placeholder: 'https://.../SitePages/page.aspx' })
          ]
        }]
      }]
    };
  }
}
