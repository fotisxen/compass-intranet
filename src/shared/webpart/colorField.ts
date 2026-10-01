// Deep import (not the package barrel) — the barrel re-exports every control
// in the library, which bloats every bundle that imports this by ~9MB.
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import type { IPropertyPaneField } from '@microsoft/sp-property-pane';

/**
 * A "dynamic color" property pane field, shared across the standalone
 * content web parts — wraps PropertyFieldColorPicker with the plumbing
 * (onPropertyChange/properties/key) every call site would otherwise repeat.
 * `webPartProperties` is the web part's own `this.properties`, `onPropertyChange`
 * is `this.onPropertyPaneFieldChanged.bind(this)`.
 */
export function colorField(
  targetProperty: string,
  label: string,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  webPartProperties: any,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
): IPropertyPaneField<any> {
  return PropertyFieldColorPicker(targetProperty, {
    key: `${targetProperty}Field`,
    label,
    selectedColor: webPartProperties[targetProperty],
    properties: webPartProperties,
    onPropertyChange,
    style: PropertyFieldColorPickerStyle.Full,
    alphaSliderHidden: true
  });
}
