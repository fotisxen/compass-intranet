import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneSlider, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import SectionHeading, { type ISectionHeadingProps } from './components/SectionHeading';

export type ISectionHeadingWebPartProps = Omit<ISectionHeadingProps, 'isEditMode'>;

export default class SectionHeadingWebPart extends BaseClientSideWebPart<ISectionHeadingWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.title = this.properties.title ?? 'Coming up';
    this.properties.linkText = this.properties.linkText ?? 'See all';
    this.properties.linkUrl = this.properties.linkUrl ?? '';
    this.properties.showArrow = this.properties.showArrow ?? true;
    this.properties.newTab = this.properties.newTab ?? false;
    this.properties.titleSize = this.properties.titleSize ?? 24;
    return super.onInit();
  }

  public render(): void {
    ReactDom.render(React.createElement(SectionHeading, { ...this.properties, isEditMode: this.displayMode === DisplayMode.Edit }), this.domElement);
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
        header: { description: 'A heading on the left and, optionally, a link on the right.' },
        groups: [{
          groupName: 'Heading',
          groupFields: [
            PropertyPaneTextField('title', { label: 'Heading' }),
            PropertyPaneSlider('titleSize', { label: 'Heading size (px at the 1440px page width)', min: 14, max: 64, step: 1, value: this.properties.titleSize ?? 24 }),
            PropertyPaneTextField('linkText', { label: 'Link text (leave empty for no link)' }),
            PropertyPaneTextField('linkUrl', { label: 'Link address — where the link goes', description: 'Paste any address: a page of this site (/sites/Intranet/SitePages/Events.aspx) or a full https:// address. The link only shows once this is filled in.', placeholder: '/sites/Intranet/SitePages/Events.aspx' }),
            PropertyPaneToggle('showArrow', { label: 'Arrow after the link text', onText: 'Yes', offText: 'No' }),
            PropertyPaneToggle('newTab', { label: 'Open the link in a new tab', onText: 'Yes', offText: 'No' })
          ]
        }]
      }]
    };
  }
}
