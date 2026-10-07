import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { SPPermission } from '@microsoft/sp-page-context';

import ChromeRoot from './components/ChromeRoot';

export interface ICompassChromeApplicationCustomizerProperties {
  chatApiUrl?: string;
  chatResourceUri?: string;
  chatQueryUrl?: string;
  /** SharePoint list holding the site navigation. Defaults to "Compass Navigation". */
  navigationList?: string;
  /** Set to false to stop the header creating the navigation list (filled with the built-in menu) when it is missing. Default true. */
  autoCreateNavigationList?: boolean;
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
            chatQueryUrl: this.properties.chatQueryUrl,
            homeUrl: this.context.pageContext.web.absoluteUrl,
            spHttpClient: this.context.spHttpClient,
            siteUrl: this.context.pageContext.web.absoluteUrl,
            navigationListTitle: this.properties.navigationList,
            autoCreateNavigationList: this.properties.autoCreateNavigationList,
            canManageLists: this.context.pageContext.web.permissions.hasPermission(SPPermission.manageLists)
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
