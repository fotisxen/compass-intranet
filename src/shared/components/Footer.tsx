import * as React from 'react';
import styles from './Footer.module.scss';
import starbulkIcon from './assets/footerLinks/starbulk.png';
import linkedinIcon from './assets/footerLinks/linkedin.webp';
import sapIcon from './assets/footerLinks/sap.webp';
import generaliIcon from './assets/footerLinks/generali.png';
import whistleblowingIcon from './assets/footerLinks/whistleblowing.png';

export interface IFooterProps {
  companyName: string;
  /** Real quote endpoint (GetStockPrice) on the same Azure Function App as the fleet API. Tickers stay hidden when absent or unreachable — no placeholder numbers are shown. */
  stockApiUrl?: string;
  /** Server-side proxy (api/src/functions/euronextStock.ts) for starbulk.com's own Euronext Athens feed — that endpoint has no CORS header, so it can't be called directly from this site. Same "stays hidden if unreachable" behavior as stockApiUrl. */
  euronextApiUrl?: string;
  /** Shifts Stay Connected/Quick Access/the ticker right (or left, negative) without moving the blue background — editable from the web part's property pane since the real page's exact alignment can't be verified until it's live. */
  contentOffsetX?: number;
}

export interface IFooterState {
  tickers: IStockTicker[];
}

interface IFooterLink {
  label: string;
  href: string;
  icon: string;
}

const STAY_CONNECTED: IFooterLink[] = [
  { label: 'Star Bulk Website', href: 'https://www.starbulk.com/', icon: starbulkIcon },
  { label: 'Star Bulk LinkedIn', href: 'https://www.linkedin.com/company/star-bulk/posts/?feedView=all', icon: linkedinIcon }
];

const QUICK_ACCESS: IFooterLink[] = [
  {
    label: 'SAP HRMS',
    href:
      'https://ay1tp8amk.accounts.cloud.sap/saml2/idp/sso/ay1tp8amk.accounts.ondemand.com?SAMLRequest=nZJNb9swDIb%2FiqC7ZCt1PUeIU2QLigXYR7C6PfQysLSyCLMlT5SW9d%2FPddKiA9YeehOol%2BTLh1xc%2FOk79tsEst7VXMmcM%2BPQt9b9qPl1cykqfrFcEPTdbNCrFPfum%2FmVDEU2JjrSx5%2Bap%2BC0B7KkHfSGdER9tfr8Sc9krofgo0ffcbYeE62DODXbxziQzjK4V3GooP8pAdEnF0li51MrCYZsqp%2FZdnyR%2F5%2FUu9b04FqJvufs0gc0k82a76Ajw9lmXfPvxdkdzstzFCXmlSgAlKjOTSHO5qjad2OwxGqUEiWzcRTBxZrP8lkp8rlQZaNyrQpdVLLK1S1n29NA7607gnpt%2BrujiPTHptmK7derhrObR%2BCjgJ%2Fw6ql7eM719cJAZMIDSr58RHk4HCQlREO0A4w%2B0AOWjPZ2GJeAKawX2fNmT5v9MlbfrLe%2Bs3jPVl3nDx%2BCgWhqHkMyE9ce4st%2BlFRTxLZiN0l1cjQYtDtrWv6WW8mWJ6v%2Fnt3yLw%3D%3D&RelayState=%2Fsf%2Fhome%3Fbplte_company%3DshipprocurD&SigAlg=http%3A%2F%2Fwww.w3.org%2F2001%2F04%2Fxmldsig-more%23rsa-sha256&Signature=ETlhPFpmM33dEnrll6%2FKsdvODNWmRgKjInYCAh2ylApOv2%2FuxTprK00TrDywbKnTIduhj%2BIhylhaLAnDyBAIrXoqUFwxxDvQKw9rEj0ZAePkEVnoUBAgf5jn9PbfZOw1ONudrTX%2BJ2izIqJ7nkMYjnWgSV1MmTPmRqknt7sv%2FSyWSJt0XhvBwD7oT9zh6rAnI7ld1k6pR0jMpmQ84TIBVz3eYikmBxLBcL%2FwAF7ZrDBoyf6uTdToua46V77p5JMoYaJ5fh4IqXC9%2FW1jSSVuiNooly0ksc8Dlh7aEaa16T2lprH2sUPch9zO3Kq5aoeGTUW8nhYufv6PKbwSPyVlvQEXeqCoo%2F2Ystsj0EK9Puiba2WNvTa%2FHNFktz2YKR2smNfrPMnpxb6k1CKBe%2ByZYTkt0LC%2FRUD4qMQDhpj%2B8RazB3Zh%2FzpFmOnW32WB2tKeZUh0wu1HLIP3MtxsVG7zy9GvbZovpdqyh6njLbjzOAaz3AuPbR3pOcYcwd8vuFne',
    icon: sapIcon
  },
  { label: 'My.generali.gr', href: 'https://my.generali.gr/', icon: generaliIcon },
  { label: 'Whistleblowing Software', href: 'https://whistleblowersoftware.com/secure/Starbulk', icon: whistleblowingIcon }
];

interface IStockTicker {
  symbol: string;
  changePct: string;
  price: string;
}

interface INasdaqApiEntry {
  symbol: string;
  lastTrade: number;
  changePercent: number;
  isDefault: string;
}

interface INasdaqApiResponse {
  data?: INasdaqApiEntry[];
}

// starbulk.com's own feed — same shape confirmed by calling
// ?module=investorrelations&action=get-euronext-stock-feed directly, but
// unlike the NASDAQ feed above, its numeric fields come through as strings.
interface IEuronextApiResponse {
  lastTrade?: string;
  changePercent?: string;
}

function formatPrice(lastTrade: number, currencySymbol: string): string {
  return `${currencySymbol}${lastTrade.toFixed(2)}`;
}

function formatChangePct(pct: number): string {
  const arrow = pct >= 0 ? '↑' : '↓';
  const sign = pct >= 0 ? '+' : '';
  return `${sign}${pct.toFixed(1).replace('.', ',')}% ${arrow}`;
}

export default class Footer extends React.Component<IFooterProps, IFooterState> {
  constructor(props: IFooterProps) {
    super(props);
    this.state = { tickers: [] };
  }

  public componentDidMount(): void {
    this._loadTickers().catch(() => {
      // Real quotes unavailable — leave tickers empty rather than show
      // placeholder numbers that would look like a working live feed.
    });
  }

  private async _loadTickers(): Promise<void> {
    // Independent try/catch per exchange (not a shared one around
    // Promise.all) — one feed being unreachable shouldn't also hide the
    // other, working one.
    const [nasdaq, euronext] = await Promise.all([
      this._loadNasdaqTicker().catch(() => undefined),
      this._loadEuronextTicker().catch(() => undefined)
    ]);

    const tickers = [nasdaq, euronext].filter((t): t is IStockTicker => !!t);
    if (tickers.length > 0) {
      this.setState({ tickers });
    }
  }

  private async _loadNasdaqTicker(): Promise<IStockTicker | undefined> {
    const { stockApiUrl } = this.props;
    if (!stockApiUrl) {
      return undefined;
    }

    const response = await fetch(stockApiUrl);
    if (!response.ok) {
      return undefined;
    }

    const data: INasdaqApiResponse = await response.json();
    const sblk = data.data?.find(e => e.isDefault === 'true');
    if (!sblk) {
      return undefined;
    }

    return { symbol: 'NASDAQ:SBLK', changePct: formatChangePct(sblk.changePercent), price: formatPrice(sblk.lastTrade, '$') };
  }

  private async _loadEuronextTicker(): Promise<IStockTicker | undefined> {
    const { euronextApiUrl } = this.props;
    if (!euronextApiUrl) {
      return undefined;
    }

    const response = await fetch(euronextApiUrl);
    if (!response.ok) {
      return undefined;
    }

    const data: IEuronextApiResponse = await response.json();
    const lastTrade = parseFloat(data.lastTrade || '');
    const changePercent = parseFloat(data.changePercent || '');
    if (isNaN(lastTrade) || isNaN(changePercent)) {
      return undefined;
    }

    return { symbol: 'EURONEXT:SBLK', changePct: formatChangePct(changePercent), price: formatPrice(lastTrade, '€') };
  }

  public render(): React.ReactElement<IFooterProps> {
    const { tickers } = this.state;
    const { contentOffsetX } = this.props;

    return (
      <footer className={styles.footer}>
        <div className={styles.outer}>
        <div className={styles.inner} style={contentOffsetX ? { marginLeft: contentOffsetX } : undefined}>
          <div className={styles.linksGroup}>
            <div className={styles.column}>
              <span className={styles.columnLabel}>STAY CONNECTED</span>
              <div className={styles.iconRow}>
                {STAY_CONNECTED.map(item => (
                  <a
                    className={styles.iconLink}
                    key={item.label}
                    href={item.href}
                    target="_blank"
                    rel="noopener noreferrer"
                    title={item.label}
                  >
                    <img src={item.icon} alt={item.label} />
                  </a>
                ))}
              </div>
            </div>

            <div className={styles.column}>
              <span className={styles.columnLabel}>QUICK ACCESS</span>
              <div className={styles.iconRow}>
                {QUICK_ACCESS.map(item => (
                  <a
                    className={styles.iconLink}
                    key={item.label}
                    href={item.href}
                    target="_blank"
                    rel="noopener noreferrer"
                    title={item.label}
                  >
                    <img src={item.icon} alt={item.label} />
                  </a>
                ))}
              </div>
            </div>
          </div>

          {tickers.length > 0 && (
            <div className={styles.tickers}>
              {tickers.map(t => (
                <div className={styles.tickerRow} key={t.symbol}>
                  <span className={styles.tickerSymbol}>{t.symbol}</span>
                  <span className={styles.tickerChange}>{t.changePct}</span>
                  <span className={styles.tickerPrice}>{t.price}</span>
                </div>
              ))}
            </div>
          )}
        </div>
        </div>
      </footer>
    );
  }
}
