import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FeatureSplitWebPartStrings';
import FeatureSplit from './components/FeatureSplit';
import { colorField } from '../../shared/webpart/colorField';

export interface IFeatureSplitWebPartProps {
  title: string;
  titleColor: string;
  bodyText: string;
  bodyColor: string;
  imageUrl?: string;
}

export default class FeatureSplitWebPart extends BaseClientSideWebPart<IFeatureSplitWebPartProps> {

  protected onInit(): Promise<void> {
    // imageUrl is genuinely optional (no default) — the other fields all
    // require a value, so they're defaulted individually rather than via
    // the Partial<...>-keyed-loop pattern used elsewhere, which doesn't
    // type-check once one property in the interface is optional.
    this.properties.title = this.properties.title ?? 'Title';
    this.properties.titleColor = this.properties.titleColor ?? '#082244';
    this.properties.bodyText = this.properties.bodyText ?? 'Introduction of each section.';
    this.properties.bodyColor = this.properties.bodyColor ?? '#082244';
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(FeatureSplit, { ...this.properties });
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
            PropertyPaneTextField('bodyText', { label: strings.BodyTextFieldLabel, multiline: true }),
            colorField('bodyColor', strings.BodyColorFieldLabel, this.properties, this._onPropertyChange),
            PropertyPaneTextField('imageUrl', { label: strings.ImageUrlFieldLabel, placeholder: 'https://.../SiteAssets/photo.jpg' })
          ]
        }]
      }]
    };
  }
}
