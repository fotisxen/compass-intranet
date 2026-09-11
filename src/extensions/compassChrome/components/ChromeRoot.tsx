import * as React from 'react';
import type { AadHttpClientFactory } from '@microsoft/sp-http';
import styles from './ChromeRoot.module.scss';
import Header from '../../../shared/components/Header';
import ChatWidget from '../../../shared/components/ChatWidget';
import { ensureInterFont } from '../../../shared/components/ensureFonts';

export interface IChromeRootProps {
  chatApiUrl: string;
  aadHttpClientFactory?: AadHttpClientFactory;
  chatResourceUri?: string;
  chatQueryUrl?: string;
}

export interface IChromeRootState {
  isChatOpen: boolean;
}

export default class ChromeRoot extends React.Component<IChromeRootProps, IChromeRootState> {
  constructor(props: IChromeRootProps) {
    super(props);
    this.state = { isChatOpen: false };
  }

  public componentDidMount(): void {
    ensureInterFont();
  }

  private _openChat = (): void => {
    this.setState({ isChatOpen: true });
  };

  private _closeChat = (): void => {
    this.setState({ isChatOpen: false });
  };

  public render(): React.ReactElement<IChromeRootProps> {
    const { isChatOpen } = this.state;

    return (
      <>
        <Header onOpenAssistant={this._openChat} />
        {isChatOpen && (
          <>
            <div className={styles.backdrop} onClick={this._closeChat} />
            <div className={styles.chatPanel}>
              <button className={styles.closeButton} onClick={this._closeChat} aria-label="Close assistant chat">✕</button>
              <ChatWidget
                apiUrl={this.props.chatApiUrl}
                aadHttpClientFactory={this.props.aadHttpClientFactory}
                chatResourceUri={this.props.chatResourceUri}
                chatQueryUrl={this.props.chatQueryUrl}
              />
            </div>
          </>
        )}
      </>
    );
  }
}
