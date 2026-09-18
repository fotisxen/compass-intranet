import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneSlider } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HolidaysWebPartStrings';
import PublicHolidays from '../../shared/components/PublicHolidays';

export interface IHolidaysWebPartProps {
  offsetX?: number;
}

export default class HolidaysWebPart extends BaseClientSideWebPart<IHolidaysWebPartProps> {

  protected onInit(): Promise<void> {
    // Backfills an instance placed before this property existed. 50 is a
    // starting estimate (not yet confirmed against real data) — ?? (not ||)
    // so a deliberately-set 0 isn't overwritten back to it.
    this.properties.offsetX = this.properties.offsetX ?? 50;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(PublicHolidays, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      offsetX: this.properties.offsetX
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
              groupName: strings.AlignmentGroupName,
              groupFields: [
                PropertyPaneSlider('offsetX', {
                  label: strings.OffsetXFieldLabel,
                  min: -300,
                  max: 300,
                  step: 5,
                  value: this.properties.offsetX
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
