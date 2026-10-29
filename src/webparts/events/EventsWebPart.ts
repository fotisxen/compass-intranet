import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EventsWebPartStrings';
import EventsWidget from '../../shared/components/EventsWidget';

export interface IEventsWebPartProps {
  eventPageUrl?: string;
}

export default class EventsWebPart extends BaseClientSideWebPart<IEventsWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.eventPageUrl = this.properties.eventPageUrl ?? `${this.context.pageContext.web.absoluteUrl}/SitePages/Event.aspx`;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(EventsWidget, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      eventPageUrl: this.properties.eventPageUrl || undefined
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
          groups: [{
            groupName: 'Event page',
            groupFields: [
              PropertyPaneTextField('eventPageUrl', { label: 'Event page address', description: '"Read more" opens this page with ?eventId=<event> added. It is the page with the Event Image and Event Details web parts.', placeholder: '/sites/Intranet/SitePages/Event.aspx' })
            ]
          }]
        }
      ]
    };
  }
}
