import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneSlider } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FooterWebPartStrings';
import Footer from '../../shared/components/Footer';

export interface IFooterWebPartProps {
  companyName: string;
  stockApiUrl?: string;
  contentOffsetX?: number;
}

export default class FooterWebPart extends BaseClientSideWebPart<IFooterWebPartProps> {

  protected onInit(): Promise<void> {
    // Backfills an instance placed before this property existed. 100 is the
    // offset confirmed correct on the real page at the time this was added
    // — ?? (not ||) so a deliberately-set 0 isn't overwritten back to it.
    this.properties.contentOffsetX = this.properties.contentOffsetX ?? 100;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(Footer, {
      companyName: this.properties.companyName || 'Compass',
      stockApiUrl: this.properties.stockApiUrl,
      contentOffsetX: this.properties.contentOffsetX
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
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('companyName', {
                  label: strings.CompanyNameFieldLabel
                }),
                PropertyPaneTextField('stockApiUrl', {
                  label: strings.StockApiUrlFieldLabel
                })
              ]
            },
            {
              groupName: strings.AlignmentGroupName,
              groupFields: [
                PropertyPaneSlider('contentOffsetX', {
                  label: strings.ContentOffsetXFieldLabel,
                  min: -300,
                  max: 300,
                  step: 5,
                  value: this.properties.contentOffsetX
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
