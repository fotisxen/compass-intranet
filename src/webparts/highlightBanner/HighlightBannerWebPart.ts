import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
// Deep import (not the package barrel) — the barrel re-exports every control
// in the library, which bloats this bundle by ~9MB for just one control.
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'HighlightBannerWebPartStrings';
import HighlightBanner from './components/HighlightBanner';
import { colorField } from '../../shared/webpart/colorField';

interface IBulletItem {
  text: string;
}

export interface IHighlightBannerWebPartProps {
  title: string;
  titleColor: string;
  items: IBulletItem[];
  bodyColor: string;
  bgColor: string;
}

const DEFAULT_ITEMS: IBulletItem[] = [
  { text: 'Coverage starts after your first 6 months of employment.' },
  { text: 'Dependents can be added after 1 year of employment.' }
];

export default class HighlightBannerWebPart extends BaseClientSideWebPart<IHighlightBannerWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.title = this.properties.title ?? 'Eligibility';
    this.properties.titleColor = this.properties.titleColor ?? '#ffffff';
    this.properties.items = this.properties.items ?? DEFAULT_ITEMS;
    this.properties.bodyColor = this.properties.bodyColor ?? '#000000';
    this.properties.bgColor = this.properties.bgColor ?? '#d9d9d9';
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(HighlightBanner, {
      title: this.properties.title,
      titleColor: this.properties.titleColor,
      items: (this.properties.items || []).map(i => i.text),
      bodyColor: this.properties.bodyColor,
      bgColor: this.properties.bgColor
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

  private _onPropertyChange = (propertyPath: string, oldValue: unknown, newValue: unknown): void => {
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();
    this.context.propertyPane.refresh();
  };

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: strings.PropertyPaneDescription },
        groups: [{
          groupName: strings.BasicGroupName,
          groupFields: [
            PropertyPaneTextField('title', { label: strings.TitleFieldLabel }),
            colorField('titleColor', strings.TitleColorFieldLabel, this.properties, this._onPropertyChange),
            PropertyFieldCollectionData('items', {
              key: 'items',
              label: strings.ItemsFieldLabel,
              panelHeader: 'Bullet points',
              manageBtnLabel: 'Manage bullet points',
              saveAndAddBtnLabel: 'Save and add another',
              fields: [
                { id: 'text', title: 'Text', type: CustomCollectionFieldType.string, required: true }
              ],
              value: this.properties.items
            }),
            colorField('bodyColor', strings.BodyColorFieldLabel, this.properties, this._onPropertyChange),
            colorField('bgColor', strings.BgColorFieldLabel, this.properties, this._onPropertyChange)
          ]
        }]
      }]
    };
  }
}
