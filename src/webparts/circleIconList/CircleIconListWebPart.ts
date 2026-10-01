import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
// Deep import (not the package barrel) — the barrel re-exports every control
// in the library, which bloats this bundle by ~9MB for just one control.
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'CircleIconListWebPartStrings';
import CircleIconList from './components/CircleIconList';

interface IListItem {
  text: string;
}

export interface ICircleIconListWebPartProps {
  circleText?: string;
  circleImageUrl?: string;
  title: string;
  items: IListItem[];
}

const DEFAULT_ITEMS: IListItem[] = [
  { text: 'List text 1' },
  { text: 'List text 2' },
  { text: 'List text 3' },
  { text: 'List text 4' },
  { text: 'List text 5' }
];

export default class CircleIconListWebPart extends BaseClientSideWebPart<ICircleIconListWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.circleText = this.properties.circleText ?? '1';
    this.properties.title = this.properties.title ?? 'Text 1';
    this.properties.items = this.properties.items ?? DEFAULT_ITEMS;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(CircleIconList, {
      circleText: this.properties.circleText,
      circleImageUrl: this.properties.circleImageUrl,
      title: this.properties.title,
      items: (this.properties.items || []).map(i => i.text)
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
            PropertyPaneTextField('circleText', { label: strings.CircleTextFieldLabel }),
            PropertyPaneTextField('circleImageUrl', { label: strings.CircleImageUrlFieldLabel, placeholder: 'https://.../SiteAssets/photo.jpg (overrides the text above)' }),
            PropertyPaneTextField('title', { label: strings.TitleFieldLabel }),
            PropertyFieldCollectionData('items', {
              key: 'items',
              label: strings.ItemsFieldLabel,
              panelHeader: 'List items',
              manageBtnLabel: 'Manage list items',
              saveAndAddBtnLabel: 'Save and add another',
              fields: [
                { id: 'text', title: 'Text', type: CustomCollectionFieldType.string, required: true }
              ],
              value: this.properties.items
            })
          ]
        }]
      }]
    };
  }
}
