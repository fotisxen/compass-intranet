import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import BackgroundBox, { type IBackgroundBoxProps } from './components/BackgroundBox';
import { colorField } from '../../shared/webpart/colorField';

export type IBackgroundBoxWebPartProps = Omit<IBackgroundBoxProps, 'isEditMode'>;

export default class BackgroundBoxWebPart extends BaseClientSideWebPart<IBackgroundBoxWebPartProps> {

  protected onInit(): Promise<void> {
    const p = this.properties;
    p.widthMode = p.widthMode ?? 'column';
    p.widthValue = p.widthValue ?? 600;
    p.widthUnit = p.widthUnit ?? 'px';
    p.align = p.align ?? 'left';
    p.heightValue = p.heightValue ?? 300;
    p.heightUnit = p.heightUnit ?? 'px';
    p.color = p.color ?? '#082244';
    p.useGradient = p.useGradient ?? false;
    p.colorEnd = p.colorEnd ?? '#123963';
    p.gradientAngle = p.gradientAngle ?? 135;
    p.opacity = p.opacity ?? 100;
    p.radius = p.radius ?? 30;
    p.perCorner = p.perCorner ?? false;
    p.radiusTopLeft = p.radiusTopLeft ?? 30;
    p.radiusTopRight = p.radiusTopRight ?? 30;
    p.radiusBottomRight = p.radiusBottomRight ?? 30;
    p.radiusBottomLeft = p.radiusBottomLeft ?? 30;
    p.borderWidth = p.borderWidth ?? 0;
    p.borderColor = p.borderColor ?? '#082244';
    p.shadow = p.shadow ?? false;
    p.pullUp = p.pullUp ?? 0;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(BackgroundBox, {
      ...this.properties,
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

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _onPropertyChange = (propertyPath: string, oldValue: any, newValue: any): void => {
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  };

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const p = this.properties;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sizeFields: IPropertyPaneField<any>[] = [
      PropertyPaneDropdown('widthMode', {
        label: 'Width',
        options: [
          { key: 'column', text: 'Fill the column' },
          { key: 'columnBleed', text: 'Fill the column, edge to edge (ignores the side padding)' },
          { key: 'full', text: 'Full page width' },
          { key: 'custom', text: 'Custom width' }
        ]
      })
    ];
    if (p.widthMode === 'custom') {
      sizeFields.push(
        PropertyPaneDropdown('widthUnit', { label: 'Width unit', options: [{ key: 'px', text: 'Pixels (px)' }, { key: '%', text: 'Percent of the column (%)' }] }),
        p.widthUnit === '%'
          ? PropertyPaneSlider('widthValue', { label: 'Width (%)', min: 5, max: 100, step: 1, value: Math.min(100, p.widthValue ?? 50) })
          : PropertyPaneSlider('widthValue', { label: 'Width (px)', min: 50, max: 2000, step: 10, value: p.widthValue ?? 600 }),
        PropertyPaneDropdown('align', { label: 'Position in the column', options: [{ key: 'left', text: 'Left' }, { key: 'center', text: 'Centre' }, { key: 'right', text: 'Right' }] })
      );
    }
    sizeFields.push(
      PropertyPaneDropdown('heightUnit', { label: 'Height unit', options: [{ key: 'px', text: 'Pixels (px)' }, { key: 'vh', text: 'Percent of the screen height (vh)' }] }),
      p.heightUnit === 'vh'
        ? PropertyPaneSlider('heightValue', { label: 'Height (vh)', min: 5, max: 100, step: 1, value: Math.min(100, p.heightValue ?? 40) })
        : PropertyPaneSlider('heightValue', { label: 'Height (px)', min: 10, max: 2000, step: 10, value: p.heightValue ?? 300 })
    );

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const colorFields: IPropertyPaneField<any>[] = [
      colorField('color', p.useGradient ? 'Start colour' : 'Colour', p, this._onPropertyChange),
      PropertyPaneToggle('useGradient', { label: 'Gradient', onText: 'On', offText: 'Off' })
    ];
    if (p.useGradient) {
      colorFields.push(
        colorField('colorEnd', 'End colour', p, this._onPropertyChange),
        PropertyPaneSlider('gradientAngle', { label: 'Gradient angle (°)', min: 0, max: 360, step: 5, value: p.gradientAngle ?? 135 })
      );
    }
    colorFields.push(PropertyPaneSlider('opacity', { label: 'Opacity (%)', min: 0, max: 100, step: 5, value: p.opacity ?? 100 }));

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const shapeFields: IPropertyPaneField<any>[] = [
      PropertyPaneToggle('perCorner', { label: 'Different radius per corner', onText: 'Yes', offText: 'No' })
    ];
    if (p.perCorner) {
      shapeFields.push(
        PropertyPaneSlider('radiusTopLeft', { label: 'Top left (px)', min: 0, max: 300, step: 1, value: p.radiusTopLeft ?? 30 }),
        PropertyPaneSlider('radiusTopRight', { label: 'Top right (px)', min: 0, max: 300, step: 1, value: p.radiusTopRight ?? 30 }),
        PropertyPaneSlider('radiusBottomRight', { label: 'Bottom right (px)', min: 0, max: 300, step: 1, value: p.radiusBottomRight ?? 30 }),
        PropertyPaneSlider('radiusBottomLeft', { label: 'Bottom left (px)', min: 0, max: 300, step: 1, value: p.radiusBottomLeft ?? 30 })
      );
    } else {
      shapeFields.push(PropertyPaneSlider('radius', { label: 'Corner radius (px)', min: 0, max: 300, step: 1, value: p.radius ?? 30 }));
    }
    shapeFields.push(
      PropertyPaneSlider('borderWidth', { label: 'Border width (px)', min: 0, max: 20, step: 1, value: p.borderWidth ?? 0 })
    );
    if ((p.borderWidth ?? 0) > 0) {
      shapeFields.push(colorField('borderColor', 'Border colour', p, this._onPropertyChange));
    }
    shapeFields.push(PropertyPaneToggle('shadow', { label: 'Shadow', onText: 'On', offText: 'Off' }));

    return {
      pages: [{
        header: { description: 'A coloured box to use as a background. SharePoint cannot put web parts inside it, so to lay content over it use "Pull the content below up" under Position.' },
        groups: [
          { groupName: 'Size', groupFields: sizeFields },
          { groupName: 'Colour', groupFields: colorFields },
          { groupName: 'Corners, border and shadow', groupFields: shapeFields },
          {
            groupName: 'Position',
            groupFields: [
              PropertyPaneSlider('pullUp', {
                label: 'Pull the content below up (px)',
                min: 0,
                max: 2000,
                step: 10,
                value: p.pullUp ?? 0
              })
            ]
          }
        ]
      }]
    };
  }
}
