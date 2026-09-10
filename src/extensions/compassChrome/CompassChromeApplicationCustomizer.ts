import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import ChromeRoot from './components/ChromeRoot';
import Footer from '../../shared/components/Footer';

export interface ICompassChromeApplicationCustomizerProperties {
  companyName?: string;
  chatApiUrl?: string;
}

export default class CompassChromeApplicationCustomizer
  extends BaseApplicationCustomizer<ICompassChromeApplicationCustomizerProperties> {

  private _topPlaceholder?: PlaceholderContent;
  private _bottomPlaceholder?: PlaceholderContent;

  @override
  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    this._renderPlaceHolders();
    return Promise.resolve();
  }

  private _renderPlaceHolders = (): void => {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
      if (this._topPlaceholder && this._topPlaceholder.domElement) {
        ReactDom.render(
          React.createElement(ChromeRoot, { chatApiUrl: this.properties.chatApiUrl || '' }),
          this._topPlaceholder.domElement
        );
      }
    }

    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);
      if (this._bottomPlaceholder && this._bottomPlaceholder.domElement) {
        ReactDom.render(
          React.createElement(Footer, { companyName: this.properties.companyName || 'Compass' }),
          this._bottomPlaceholder.domElement
        );
      }
    }
  };

  @override
  public onDispose(): void {
    if (this._topPlaceholder && this._topPlaceholder.domElement) {
      ReactDom.unmountComponentAtNode(this._topPlaceholder.domElement);
    }
    if (this._bottomPlaceholder && this._bottomPlaceholder.domElement) {
      ReactDom.unmountComponentAtNode(this._bottomPlaceholder.domElement);
    }
    super.onDispose();
  }
}
