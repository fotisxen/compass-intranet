import * as React from 'react';
import type { AadHttpClientFactory, SPHttpClient } from '@microsoft/sp-http';
import styles from './ChromeRoot.module.scss';
import Header from '../../../shared/components/Header';
import ChatWidget from '../../../shared/components/ChatWidget';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import { NAV_ITEMS, type INavItem } from '../../../shared/components/navData';
import { NAV_LIST_TITLE, ensureNavigationList, loadNavFromList } from '../../../shared/components/navigationList';

export interface IChromeRootProps {
  chatApiUrl: string;
  aadHttpClientFactory?: AadHttpClientFactory;
  chatResourceUri?: string;
  chatQueryUrl?: string;
  homeUrl?: string;
  /** Needed to read the navigation list. Without them the built-in menu is shown. */
  spHttpClient?: SPHttpClient;
  siteUrl?: string;
  /** Title of the navigation list. Defaults to "Compass Navigation". */
  navigationListTitle?: string;
  /** Whether the signed-in user can create lists — only then is the list created (and filled with the built-in menu) when it's missing. */
  canManageLists?: boolean;
  autoCreateNavigationList?: boolean;
}

export interface IChromeRootState {
  isChatOpen: boolean;
  navItems: INavItem[];
}

function cacheKey(siteUrl: string): string {
  return `compass.nav.v1.${siteUrl.toLowerCase()}`;
}

// The last menu read from the list is remembered, so the next page opens with
// the right menu instantly instead of flashing the built-in one first.
function readCachedNav(siteUrl?: string): INavItem[] | undefined {
  if (!siteUrl) {
    return undefined;
  }
  try {
    const raw = window.localStorage.getItem(cacheKey(siteUrl));
    const parsed: unknown = raw ? JSON.parse(raw) : undefined;
    return Array.isArray(parsed) && parsed.length > 0 ? (parsed as INavItem[]) : undefined;
  } catch {
    return undefined;
  }
}

function writeCachedNav(siteUrl: string, items: INavItem[]): void {
  try {
    window.localStorage.setItem(cacheKey(siteUrl), JSON.stringify(items));
  } catch {
    // Storage unavailable (private mode, quota) — the menu still works, just without the cache.
  }
}

export default class ChromeRoot extends React.Component<IChromeRootProps, IChromeRootState> {
  constructor(props: IChromeRootProps) {
    super(props);
    this.state = { isChatOpen: false, navItems: readCachedNav(props.siteUrl) || NAV_ITEMS };
  }

  public componentDidMount(): void {
    ensureInterFont();
    this._loadNavigation().catch((err: unknown) => {
      // List unreachable or couldn't be set up — whatever menu is showing
      // (cached or built-in) stays; the console says why.
      console.warn('[Compass header] Navigation list unavailable, keeping the current menu:', err);
    });
  }

  private async _loadNavigation(): Promise<void> {
    const { spHttpClient, siteUrl, canManageLists, autoCreateNavigationList } = this.props;
    if (!spHttpClient || !siteUrl) {
      return;
    }
    const listTitle = this.props.navigationListTitle || NAV_LIST_TITLE;

    let result = await loadNavFromList(spHttpClient, siteUrl, listTitle);
    const needsSetup = result === 'missing' || result === 'unreadable' || (Array.isArray(result) && result.length === 0);

    // First owner visit: create the list and fill it with the current menu, so
    // nothing the site shows today is lost. Tried once per browser session.
    const attemptedKey = `compass.nav.setup.${siteUrl.toLowerCase()}`;
    if (needsSetup && canManageLists && autoCreateNavigationList !== false && !window.sessionStorage.getItem(attemptedKey)) {
      window.sessionStorage.setItem(attemptedKey, '1');
      await ensureNavigationList(spHttpClient, siteUrl, listTitle, NAV_ITEMS);
      result = await loadNavFromList(spHttpClient, siteUrl, listTitle);
    }

    if (Array.isArray(result) && result.length > 0) {
      writeCachedNav(siteUrl, result);
      this.setState({ navItems: result });
    }
  }

  private _openChat = (): void => {
    this.setState({ isChatOpen: true });
  };

  private _closeChat = (): void => {
    this.setState({ isChatOpen: false });
  };

  public render(): React.ReactElement<IChromeRootProps> {
    const { isChatOpen, navItems } = this.state;

    return (
      <>
        <Header onOpenAssistant={this._openChat} homeUrl={this.props.homeUrl} navItems={navItems} />
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
