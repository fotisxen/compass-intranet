import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
// Deep import (not the package barrel) — the barrel re-exports every control
// in the library, which bloats this bundle by ~9MB for just one control.
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'DualColumnFeatureWebPartStrings';
import DualColumnFeature from './components/DualColumnFeature';
import { colorField } from '../../shared/webpart/colorField';

interface IBulletItem {
  text: string;
}

export interface IDualColumnFeatureWebPartProps {
  bgColor: string;
  leftTitle: string;
  leftItems: IBulletItem[];
  rightTitle: string;
  rightItems: IBulletItem[];
}

const DEFAULT_LEFT_ITEMS: IBulletItem[] = [{ text: 'Introduction of each section.' }];
const DEFAULT_RIGHT_ITEMS: IBulletItem[] = [{ text: 'Introduction of each section.' }];

export default class DualColumnFeatureWebPart extends BaseClientSideWebPart<IDualColumnFeatureWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.bgColor = this.properties.bgColor ?? '#082244';
    this.properties.leftTitle = this.properties.leftTitle ?? 'Title 1';
    this.properties.leftItems = this.properties.leftItems ?? DEFAULT_LEFT_ITEMS;
    this.properties.rightTitle = this.properties.rightTitle ?? 'Title 2';
    this.properties.rightItems = this.properties.rightItems ?? DEFAULT_RIGHT_ITEMS;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(DualColumnFeature, {
      bgColor: this.properties.bgColor,
      leftTitle: this.properties.leftTitle,
      leftItems: (this.properties.leftItems || []).map(i => i.text),
      rightTitle: this.properties.rightTitle,
      rightItems: (this.properties.rightItems || []).map(i => i.text)
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
            colorField('bgColor', strings.BgColorFieldLabel, this.properties, this._onPropertyChange),
            PropertyPaneTextField('leftTitle', { label: strings.LeftTitleFieldLabel }),
            PropertyFieldCollectionData('leftItems', {
              key: 'leftItems',
              label: strings.LeftItemsFieldLabel,
              panelHeader: 'Left card bullet points',
              manageBtnLabel: 'Manage bullet points',
              saveAndAddBtnLabel: 'Save and add another',
              fields: [
                { id: 'text', title: 'Text', type: CustomCollectionFieldType.string, required: true }
              ],
              value: this.properties.leftItems
            }),
            PropertyPaneTextField('rightTitle', { label: strings.RightTitleFieldLabel }),
            PropertyFieldCollectionData('rightItems', {
              key: 'rightItems',
              label: strings.RightItemsFieldLabel,
              panelHeader: 'Right card bullet points',
              manageBtnLabel: 'Manage bullet points',
              saveAndAddBtnLabel: 'Save and add another',
              fields: [
                { id: 'text', title: 'Text', type: CustomCollectionFieldType.string, required: true }
              ],
              value: this.properties.rightItems
            })
          ]
        }]
      }]
    };
  }
}
