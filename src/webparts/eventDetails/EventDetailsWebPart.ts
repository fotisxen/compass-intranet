import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneSlider, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import EventDetails from './components/EventDetails';
import { DEFAULT_EVENTS_LIST } from '../../shared/events/eventsApi';

export interface IEventDetailsWebPartProps {
  listRef?: string;
  showTitle?: boolean;
  titleSize?: number;
  showWhen?: boolean;
  showTimezone?: boolean;
  showCategory?: boolean;
  showLocation?: boolean;
  showAbout?: boolean;
  spaceLeft?: number;
}

export default class EventDetailsWebPart extends BaseClientSideWebPart<IEventDetailsWebPartProps> {

  protected onInit(): Promise<void> {
    const p = this.properties;
    p.listRef = p.listRef ?? DEFAULT_EVENTS_LIST;
    p.showTitle = p.showTitle ?? true;
    p.titleSize = p.titleSize ?? 40;
    p.showWhen = p.showWhen ?? true;
    p.showTimezone = p.showTimezone ?? true;
    p.showCategory = p.showCategory ?? true;
    p.showLocation = p.showLocation ?? true;
    p.showAbout = p.showAbout ?? true;
    p.spaceLeft = p.spaceLeft ?? 32;
    return super.onInit();
  }

  public render(): void {
    const p = this.properties;
    const element: React.ReactElement = React.createElement(EventDetails, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      listRef: p.listRef || DEFAULT_EVENTS_LIST,
      showTitle: p.showTitle !== false,
      titleSize: p.titleSize ?? 40,
      showWhen: p.showWhen !== false,
      showTimezone: p.showTimezone !== false,
      showCategory: p.showCategory !== false,
      showLocation: p.showLocation !== false,
      showAbout: p.showAbout !== false,
      spaceLeft: p.spaceLeft ?? 32,
      isEditMode: this.displayMode === DisplayMode.Edit
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
        header: { description: 'Shows the event in the page address (?eventId=). While editing, it shows the next upcoming event.' },
        groups: [
          {
            groupName: 'Layout',
            groupFields: [PropertyPaneSlider('spaceLeft', { label: 'Space on the left, between this and the image next to it (px)', min: 0, max: 200, step: 4, value: this.properties.spaceLeft ?? 32 })]
          },
          {
            groupName: 'Events list',
            groupFields: [PropertyPaneTextField('listRef', { label: 'Events list (name or id)' })]
          },
          {
            groupName: 'What to show',
            groupFields: [
              PropertyPaneToggle('showTitle', { label: 'Title', onText: 'Show', offText: 'Hide' }),
              PropertyPaneSlider('titleSize', { label: 'Title size (px at the 1440px page width)', min: 16, max: 80, step: 1, value: this.properties.titleSize ?? 40 }),
              PropertyPaneToggle('showWhen', { label: 'When', onText: 'Show', offText: 'Hide' }),
              PropertyPaneToggle('showTimezone', { label: 'Time zone line under "When"', onText: 'Show', offText: 'Hide' }),
              PropertyPaneToggle('showCategory', { label: 'Category', onText: 'Show', offText: 'Hide' }),
              PropertyPaneToggle('showLocation', { label: 'Location', onText: 'Show', offText: 'Hide' }),
              PropertyPaneToggle('showAbout', { label: 'About (the event\'s description)', onText: 'Show', offText: 'Hide' })
            ]
          }
        ]
      }]
    };
  }
}
