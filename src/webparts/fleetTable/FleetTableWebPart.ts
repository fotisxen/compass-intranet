import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { SPPermission } from '@microsoft/sp-page-context';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FleetTableWebPartStrings';
import FleetTable from './components/FleetTable';
import { FLEET_LIST_TITLE, type VesselSortField } from '../../shared/components/fleetVesselsList';

export interface IFleetTableWebPartProps {
  listTitle?: string;
  sortBy?: VesselSortField;
  sortDescending?: boolean;
  maxRows?: number;
}

export default class FleetTableWebPart extends BaseClientSideWebPart<IFleetTableWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.listTitle = this.properties.listTitle ?? FLEET_LIST_TITLE;
    this.properties.sortBy = this.properties.sortBy ?? 'dwt';
    this.properties.sortDescending = this.properties.sortDescending ?? true;
    this.properties.maxRows = this.properties.maxRows ?? 0;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(FleetTable, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      listTitle: this.properties.listTitle,
      sortBy: this.properties.sortBy || 'dwt',
      sortDescending: this.properties.sortDescending !== false,
      maxRows: this.properties.maxRows || 0,
      canManageLists: this.context.pageContext.web.permissions.hasPermission(SPPermission.manageLists)
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
      pages: [{
        header: { description: strings.PropertyPaneDescription },
        groups: [{
          groupName: strings.BasicGroupName,
          groupFields: [
            PropertyPaneTextField('listTitle', { label: strings.ListTitleFieldLabel }),
            PropertyPaneDropdown('sortBy', {
              label: strings.SortByFieldLabel,
              options: [
                { key: 'dwt', text: 'DWT' },
                { key: 'name', text: 'Vessel name' },
                { key: 'built', text: 'Year built' }
              ]
            }),
            PropertyPaneToggle('sortDescending', { label: strings.SortDescendingFieldLabel }),
            PropertyPaneSlider('maxRows', { label: strings.MaxRowsFieldLabel, min: 0, max: 100, step: 1, showValue: true })
          ]
        }]
      }]
    };
  }
}
