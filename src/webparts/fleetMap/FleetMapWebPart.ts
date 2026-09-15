import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FleetMapWebPartStrings';
import FleetMap from '../../shared/components/FleetMap';

export interface IFleetMapWebPartProps {
  // URL of the fleet-positions Azure Function proxy (api/src/functions/fleet.ts).
  // Public/non-secret — the real provider's API key lives server-side in the
  // Function App's settings, never here. Leave blank to keep office pins.
  fleetApiUrl?: string;
}

export default class FleetMapWebPart extends BaseClientSideWebPart<IFleetMapWebPartProps> {

  public render(): void {
    const element: React.ReactElement = React.createElement(FleetMap, {
      fleetApiUrl: this.properties.fleetApiUrl
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
                PropertyPaneTextField('fleetApiUrl', {
                  label: strings.FleetApiUrlFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
