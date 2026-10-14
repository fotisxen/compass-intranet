import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// Deep import (not the package barrel) — the barrel bloats the bundle by ~9MB.
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'NewsGridWebPartStrings';
import NewsGrid from './components/NewsGrid';

export interface INewsGridWebPartProps {
  pinnedNews?: { pageUrl: string }[];
}

export default class NewsGridWebPart extends BaseClientSideWebPart<INewsGridWebPartProps> {

  public render(): void {
    const element: React.ReactElement = React.createElement(NewsGrid, {
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
          header: { description: strings.PropertyPaneDescription },
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
