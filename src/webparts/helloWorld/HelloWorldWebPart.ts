import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// Deep import (not the package barrel) — the barrel re-exports every control
// in the library, which bloats this bundle by ~9MB for just one control.
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'HelloWorldWebPartStrings';
import HomeNews from './components/HomeNews';

// This is the "Compass Home News Entity" web part. It keeps its original
// HelloWorld id/alias so pages that already have it keep their pinned news;
// the people row is the separate "Compass Home People Entity" web part.
export interface IHelloWorldWebPartProps {
  pinnedNews?: { pageUrl: string }[];
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public render(): void {
    const element: React.ReactElement = React.createElement(HomeNews, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      pinnedNews: (this.properties.pinnedNews || []).map(p => p.pageUrl).filter(u => !!u && !!u.trim())
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
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: 'News',
              groupFields: [
                PropertyFieldCollectionData('pinnedNews', {
                  key: 'pinnedNews',
                  label: 'Pinned news (always visible, in this order — the number at the start of each row in the panel sets its position). The other cards rotate automatically in the remaining spots.',
                  panelHeader: 'News — pinned pages',
                  manageBtnLabel: 'Manage pinned news',
                  saveAndAddBtnLabel: 'Save and add another',
                  enableSorting: true,
                  fields: [
                    { id: 'pageUrl', title: 'Page URL or file name', type: CustomCollectionFieldType.string, required: true, placeholder: '/sites/Intranet/SitePages/My-News.aspx' }
                  ],
                  value: this.properties.pinnedNews || []
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
