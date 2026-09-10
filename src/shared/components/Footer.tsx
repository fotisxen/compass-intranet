import * as React from 'react';
import styles from './Footer.module.scss';

export interface IFooterProps {
  companyName: string;
}

const QUICK_ACCESS: string[] = ['SAP', 'DS', 'VV'];

interface IStockTicker {
  symbol: string;
  changePct: string;
  price: string;
}

const STOCK_TICKERS: IStockTicker[] = [
  { symbol: 'SBLK', changePct: '+2,0% ↑', price: '$26.56' },
  { symbol: 'STK2', changePct: '+2,2% ↑', price: '$14.56' }
];

export default class Footer extends React.Component<IFooterProps> {
  public render(): React.ReactElement<IFooterProps> {
    const { companyName } = this.props;

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

          <div className={styles.tickers}>
            {STOCK_TICKERS.map(t => (
              <div className={styles.tickerRow} key={t.symbol}>
                <span className={styles.tickerSymbol}>{t.symbol}</span>
                <span className={styles.tickerChange}>{t.changePct}</span>
                <span className={styles.tickerPrice}>{t.price}</span>
              </div>
            ))}
          </div>
        </div>
      </footer>
    );
  }
}
