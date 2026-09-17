import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import ChromeRoot from './components/ChromeRoot';

export interface ICompassChromeApplicationCustomizerProperties {
  chatApiUrl?: string;
  chatResourceUri?: string;
  chatQueryUrl?: string;
}

export default class CompassChromeApplicationCustomizer
  extends BaseApplicationCustomizer<ICompassChromeApplicationCustomizerProperties> {

  private _topPlaceholder?: PlaceholderContent;

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
          React.createElement(ChromeRoot, {
            chatApiUrl: this.properties.chatApiUrl || '',
            aadHttpClientFactory: this.context.aadHttpClientFactory,
            chatResourceUri: this.properties.chatResourceUri,
            chatQueryUrl: this.properties.chatQueryUrl
          }),
          this._topPlaceholder.domElement
        );
      }
    }
  };

  @override
  public onDispose(): void {
    if (this._topPlaceholder && this._topPlaceholder.domElement) {
      ReactDom.unmountComponentAtNode(this._topPlaceholder.domElement);
    }
    super.onDispose();
  }
}
