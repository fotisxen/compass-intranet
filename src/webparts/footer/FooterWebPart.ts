import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FooterWebPartStrings';
import Footer from '../../shared/components/Footer';

export interface IFooterWebPartProps {
  companyName: string;
  stockApiUrl?: string;
}

export default class FooterWebPart extends BaseClientSideWebPart<IFooterWebPartProps> {

  public render(): void {
    const element: React.ReactElement = React.createElement(Footer, {
      companyName: this.properties.companyName || 'Compass',
      stockApiUrl: this.properties.stockApiUrl
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
            }
          ]
        }
      ]
    };
  }
}
