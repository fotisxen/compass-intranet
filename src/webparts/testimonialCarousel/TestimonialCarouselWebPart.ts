import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
// Deep import (not the package barrel) — the barrel re-exports every control
// in the library, which bloats this bundle by ~9MB for just one control.
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'TestimonialCarouselWebPartStrings';
import TestimonialCarousel, { type ITestimonial } from './components/TestimonialCarousel';
import { colorField } from '../../shared/webpart/colorField';

export interface ITestimonialCarouselWebPartProps {
  title: string;
  titleColor: string;
  bgColor: string;
  testimonials: ITestimonial[];
}

const DEFAULT_TESTIMONIALS: ITestimonial[] = [
  { quote: 'When I donate blood I feel like I am giving life to someone in need — without expecting anything in return.', author: 'Dimitris Spyrou' },
  { quote: 'When I donate blood I feel like I am giving life to someone in need — without expecting anything in return.', author: 'Dimitris Spyrou' },
  { quote: 'When I donate blood I feel like I am giving life to someone in need — without expecting anything in return.', author: 'Dimitris Spyrou' }
];

export default class TestimonialCarouselWebPart extends BaseClientSideWebPart<ITestimonialCarouselWebPartProps> {

  protected onInit(): Promise<void> {
    this.properties.title = this.properties.title ?? 'Title';
    this.properties.titleColor = this.properties.titleColor ?? '#082244';
    this.properties.bgColor = this.properties.bgColor ?? '#e6e6e6';
    this.properties.testimonials = this.properties.testimonials ?? DEFAULT_TESTIMONIALS;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(TestimonialCarousel, { ...this.properties });
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
            colorField('bgColor', strings.BgColorFieldLabel, this.properties, this._onPropertyChange),
            PropertyFieldCollectionData('testimonials', {
              key: 'testimonials',
              label: strings.TestimonialsFieldLabel,
              panelHeader: 'Testimonials',
              manageBtnLabel: 'Manage testimonials',
              saveAndAddBtnLabel: 'Save and add another',
              fields: [
                { id: 'quote', title: 'Quote', type: CustomCollectionFieldType.string, required: true },
                { id: 'author', title: 'Author', type: CustomCollectionFieldType.string, required: true }
              ],
              value: this.properties.testimonials
            })
          ]
        }]
      }]
    };
  }
}
