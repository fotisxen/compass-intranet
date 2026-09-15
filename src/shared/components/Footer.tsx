import * as React from 'react';
import styles from './Footer.module.scss';

export interface IFooterProps {
  companyName: string;
  /** Real quote endpoint (GetStockPrice) on the same Azure Function App as the fleet API. Tickers stay hidden when absent or unreachable — no placeholder numbers are shown. */
  stockApiUrl?: string;
}

export interface IFooterState {
  tickers: IStockTicker[];
}

const QUICK_ACCESS: string[] = ['SAP', 'DS', 'VV'];

interface IStockTicker {
  symbol: string;
  changePct: string;
  price: string;
}

interface IStockApiEntry {
  symbol: string;
  lastTrade: number;
  changePercent: number;
  isDefault: string;
}

interface IStockApiResponse {
  data?: IStockApiEntry[];
}

function formatPrice(lastTrade: number): string {
  return `$${lastTrade.toFixed(2)}`;
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
    const { stockApiUrl } = this.props;
    if (!stockApiUrl) {
      return;
    }

    const response = await fetch(stockApiUrl);
    if (!response.ok) {
      return;
    }

    const data: IStockApiResponse = await response.json();
    if (!data.data || data.data.length === 0) {
      return;
    }

    const sblk = data.data.filter(e => e.isDefault === 'true')[0];
    const index = data.data.filter(e => e.symbol === '.SPX')[0];
    const tickers: IStockTicker[] = [sblk, index]
      .filter((e): e is IStockApiEntry => !!e)
      .map(e => ({ symbol: e.symbol, changePct: formatChangePct(e.changePercent), price: formatPrice(e.lastTrade) }));

    if (tickers.length > 0) {
      this.setState({ tickers });
    }
  }

  public render(): React.ReactElement<IFooterProps> {
    const { companyName } = this.props;
    const { tickers } = this.state;

    return (
      <footer className={styles.footer}>
        <div className={styles.inner}>
          <div className={styles.column}>
            <span className={styles.columnLabel}>STAY CONNECTED</span>
            <div className={styles.iconRow}>
              <span className={styles.iconCircle}>{companyName.charAt(0)}</span>
              <span className={styles.iconCircle}>in</span>
            </div>
          </div>

          <div className={styles.column}>
            <span className={styles.columnLabel}>QUICK ACCESS</span>
            <div className={styles.iconRow}>
              {QUICK_ACCESS.map(item => (
                <span className={styles.iconCircle} key={item}>{item}</span>
              ))}
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
      </footer>
    );
  }
}
