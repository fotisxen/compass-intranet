import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneSlider, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import LinkButton, { type ILinkButtonProps } from './components/LinkButton';
import { colorField } from '../../shared/webpart/colorField';

export type ILinkButtonWebPartProps = ILinkButtonProps;

export default class LinkButtonWebPart extends BaseClientSideWebPart<ILinkButtonWebPartProps> {

  protected onInit(): Promise<void> {
    const p = this.properties;
    p.label = p.label ?? 'See All Events';
    p.url = p.url ?? '';
    p.showArrow = p.showArrow ?? true;
    p.newTab = p.newTab ?? false;
    p.color = p.color ?? '#A5192D';
    p.textColor = p.textColor ?? '#FFFFFF';
    p.align = p.align ?? 'left';
    p.spaceAbove = p.spaceAbove ?? 40;
    return super.onInit();
  }

  public render(): void {
    ReactDom.render(React.createElement(LinkButton, { ...this.properties }), this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _onPropertyChange = (propertyPath: string, oldValue: any, newValue: any): void => {
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  };

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'A rounded button that links to a page.' },
        groups: [
          {
            groupName: 'Button',
            groupFields: [
              PropertyPaneTextField('label', { label: 'Text' }),
              PropertyPaneTextField('url', { label: 'Link address', placeholder: '/sites/Intranet/SitePages/Events.aspx' }),
              PropertyPaneToggle('showArrow', { label: 'Arrow after the text', onText: 'Yes', offText: 'No' }),
              PropertyPaneToggle('newTab', { label: 'Open in a new tab', onText: 'Yes', offText: 'No' }),
              PropertyPaneDropdown('align', { label: 'Position', options: [{ key: 'left', text: 'Left' }, { key: 'center', text: 'Centre' }, { key: 'right', text: 'Right' }] }),
              PropertyPaneSlider('spaceAbove', { label: 'Space above the button (px)', min: 0, max: 300, step: 4, value: this.properties.spaceAbove ?? 40 })
            ]
          },
          {
            groupName: 'Colours',
            groupFields: [
              colorField('color', 'Button colour', this.properties, this._onPropertyChange),
              colorField('textColor', 'Text colour', this.properties, this._onPropertyChange)
            ]
          }
        ]
      }]
    };
  }
}
