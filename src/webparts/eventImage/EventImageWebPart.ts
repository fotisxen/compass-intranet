import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneSlider, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import EventImage from './components/EventImage';
import { DEFAULT_EVENTS_LIST } from '../../shared/events/eventsApi';

export interface IEventImageWebPartProps {
  listRef?: string;
  aspectRatio?: string;
  radius?: number;
}

export default class EventImageWebPart extends BaseClientSideWebPart<IEventImageWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.listRef = this.properties.listRef ?? DEFAULT_EVENTS_LIST;
    this.properties.aspectRatio = this.properties.aspectRatio ?? '5 / 3';
    this.properties.radius = this.properties.radius ?? 0;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(EventImage, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      listRef: this.properties.listRef || DEFAULT_EVENTS_LIST,
      aspectRatio: this.properties.aspectRatio || '5 / 3',
      radius: this.properties.radius ?? 0,
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
        header: { description: 'Shows the image of the event in the page address (?eventId=). While editing, it shows the next upcoming event.' },
        groups: [{
          groupName: 'Image',
          groupFields: [
            PropertyPaneTextField('listRef', { label: 'Events list (name or id)' }),
            PropertyPaneDropdown('aspectRatio', {
              label: 'Shape',
              options: [
                { key: '5 / 3', text: 'Wide (5:3)' },
                { key: '16 / 9', text: 'Wide (16:9)' },
                { key: '4 / 3', text: 'Photo (4:3)' },
                { key: '1 / 1', text: 'Square' }
              ]
            }),
            PropertyPaneSlider('radius', { label: 'Corner radius (px)', min: 0, max: 80, step: 1, value: this.properties.radius ?? 0 })
          ]
        }]
      }]
    };
  }
}
