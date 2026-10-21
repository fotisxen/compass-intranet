import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneSlider, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import EventsGrid from './components/EventsGrid';
import { DEFAULT_EVENTS_LIST } from '../../shared/events/eventsApi';

export interface IEventsGridWebPartProps {
  listRef?: string;
  eventPageUrl?: string;
  maxEvents?: number;
  emptyText?: string;
}

export default class EventsGridWebPart extends BaseClientSideWebPart<IEventsGridWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.listRef = this.properties.listRef ?? DEFAULT_EVENTS_LIST;
    this.properties.eventPageUrl = this.properties.eventPageUrl ?? `${this.context.pageContext.web.absoluteUrl}/SitePages/Event.aspx`;
    this.properties.maxEvents = this.properties.maxEvents ?? 48;
    this.properties.emptyText = this.properties.emptyText ?? 'No upcoming events.';
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(EventsGrid, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      listRef: this.properties.listRef || DEFAULT_EVENTS_LIST,
      eventPageUrl: this.properties.eventPageUrl || '',
      maxEvents: this.properties.maxEvents || 48,
      emptyText: this.properties.emptyText || ''
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
        header: { description: 'The events come from the site\'s Events (calendar) list. Only events from today onwards are shown, soonest first.' },
        groups: [{
          groupName: 'Events',
          groupFields: [
            PropertyPaneTextField('listRef', { label: 'Events list (name or id)', description: 'Leave as it is to use the site\'s standard Events list.' }),
            PropertyPaneTextField('eventPageUrl', { label: 'Event page address', description: 'The page that has the Event Image and Event Details web parts. A click on a card opens it with ?id=<event> added.', placeholder: '/sites/Intranet/SitePages/Event.aspx' }),
            PropertyPaneSlider('maxEvents', { label: 'Most events to show', min: 4, max: 200, step: 4, value: this.properties.maxEvents ?? 48 }),
            PropertyPaneTextField('emptyText', { label: 'Text when there are no events' })
          ]
        }]
      }]
    };
  }
}
