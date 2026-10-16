import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { SPPermission } from '@microsoft/sp-page-context';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import MarketplaceFilters from './components/MarketplaceFilters';

export interface IMarketplaceFiltersWebPartProps {
  // No settings: the buttons come from the Compass Marketplace Categories list.
  placeholder?: string;
}

export default class MarketplaceFiltersWebPart extends BaseClientSideWebPart<IMarketplaceFiltersWebPartProps> {

  public render(): void {
    const element: React.ReactElement = React.createElement(MarketplaceFilters, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      canManageLists: this.context.pageContext.web.permissions.hasPermission(SPPermission.manageLists)
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
        header: { description: 'The category buttons come from the "Compass Marketplace Categories" list: add, rename, reorder (Sort order) or delete rows there. "All" is always shown first.' },
        groups: []
      }]
    };
  }
}
