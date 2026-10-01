import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
// Deep import (not the package barrel) — the barrel re-exports every control
// in the library, which bloats this bundle by ~9MB for just one control.
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'StatFeatureWebPartStrings';
import StatFeature, { type IStatColumn } from './components/StatFeature';
import { colorField } from '../../shared/webpart/colorField';

export interface IStatFeatureWebPartProps {
  bgColor: string;
  eyebrowText?: string;
  eyebrowColor: string;
  titleText: string;
  titleColor: string;
  bodyText: string;
  bodyColor: string;
  imageUrl?: string;
  columnsColor: string;
  columns: IStatColumn[];
}

const DEFAULT_COLUMNS: IStatColumn[] = [
  { title: 'Title 1', text: 'Introduction of each section.' },
  { title: 'Title 2', text: 'Introduction of each section.' },
  { title: 'Title 3', text: 'Introduction of each section.' }
];

export default class StatFeatureWebPart extends BaseClientSideWebPart<IStatFeatureWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.bgColor = this.properties.bgColor ?? '#082244';
    this.properties.eyebrowColor = this.properties.eyebrowColor ?? '#ffffff';
    this.properties.titleText = this.properties.titleText ?? 'Big Title';
    this.properties.titleColor = this.properties.titleColor ?? '#ffffff';
    this.properties.bodyText = this.properties.bodyText ?? 'Introduction of each section.';
    this.properties.bodyColor = this.properties.bodyColor ?? '#ffffff';
    this.properties.columnsColor = this.properties.columnsColor ?? '#ffffff';
    this.properties.columns = this.properties.columns ?? DEFAULT_COLUMNS;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(StatFeature, { ...this.properties });
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
            PropertyPaneTextField('eyebrowText', { label: strings.EyebrowTextFieldLabel }),
            colorField('eyebrowColor', strings.EyebrowColorFieldLabel, this.properties, this._onPropertyChange),
            PropertyPaneTextField('titleText', { label: strings.TitleTextFieldLabel }),
            colorField('titleColor', strings.TitleColorFieldLabel, this.properties, this._onPropertyChange),
            PropertyPaneTextField('bodyText', { label: strings.BodyTextFieldLabel, multiline: true }),
            colorField('bodyColor', strings.BodyColorFieldLabel, this.properties, this._onPropertyChange),
            PropertyPaneTextField('imageUrl', { label: strings.ImageUrlFieldLabel, placeholder: 'https://.../SiteAssets/photo.jpg (optional)' }),
            colorField('columnsColor', strings.ColumnsColorFieldLabel, this.properties, this._onPropertyChange),
            PropertyFieldCollectionData('columns', {
              key: 'columns',
              label: strings.ColumnsFieldLabel,
              panelHeader: 'Columns',
              manageBtnLabel: 'Manage columns',
              saveAndAddBtnLabel: 'Save and add another',
              fields: [
                { id: 'title', title: 'Title', type: CustomCollectionFieldType.string, required: true },
                { id: 'text', title: 'Text', type: CustomCollectionFieldType.string, required: true }
              ],
              value: this.properties.columns
            })
          ]
        }]
      }]
    };
  }
}
