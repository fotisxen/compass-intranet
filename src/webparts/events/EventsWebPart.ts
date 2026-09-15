import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EventsWebPartStrings';
import EventsWidget from '../../shared/components/EventsWidget';

export type IEventsWebPartProps = Record<string, never>;

export default class EventsWebPart extends BaseClientSideWebPart<IEventsWebPartProps> {

  public render(): void {
    const element: React.ReactElement = React.createElement(EventsWidget, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl
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
          groups: []
        }
      ]
    };
  }
}
