// One entry of the stock quotes API's `data` list (the GetStockPrice endpoint
// that fronts the investor-relations quotes feed). Only the fields the footer
// reads are typed; the real entries carry many more.
export interface IQuoteEntry {
  symbol?: string;
  exchange?: string;
  exchangeShortCode?: string;
  lastTrade?: number;
  changePercent?: number;
  isDefault?: string;
  isIndex?: string;
}

// Words that identify an entry as the NASDAQ or the Euronext Athens listing.
// The feed is known to label NASDAQ ("NASDAQ" / "NSD"); the Euronext entry
// doesn't exist yet, so it's matched on the several ways it could plausibly be
// labelled (exchange name, short code, or a ".ATH"-style symbol suffix).
const NASDAQ_LIKE = /nasdaq|\bnsd\b|\bnms\b|\bngs\b|\bngm\b/i;
const EURONEXT_LIKE = /euronext|athens|athex|\bath\b|\bxath\b|\benx\b/i;

function label(entry: IQuoteEntry): string {
  return `${entry.exchange || ''} ${entry.exchangeShortCode || ''} ${entry.symbol || ''}`;
}

/** Star Bulk's own share (SBLK), not an index, with a usable price and change. */
function sblkStocks(entries: IQuoteEntry[] | undefined): IQuoteEntry[] {
  return (entries || []).filter(e =>
    e.isIndex !== 'true'
    && /^sblk/i.test(e.symbol || '')
    && typeof e.lastTrade === 'number' && e.lastTrade > 0
    && typeof e.changePercent === 'number'
  );
}

/** The NASDAQ listing: labelled as NASDAQ, falling back to the feed's own default flag (how it was always picked). */
export function pickNasdaqEntry(entries: IQuoteEntry[] | undefined): IQuoteEntry | undefined {
  const stocks = sblkStocks(entries);
  return stocks.find(e => NASDAQ_LIKE.test(label(e)))
    || (entries || []).find(e => e.isDefault === 'true' && typeof e.lastTrade === 'number' && typeof e.changePercent === 'number');
}

/**
 * The Euronext Athens listing: an SBLK stock entry labelled as Euronext/Athens,
 * or failing that any other SBLK stock entry that isn't labelled NASDAQ.
 * Resolves to undefined while the feed has no such entry — so the footer simply
 * shows no Euronext ticker until the feed provides one.
 */
export function pickEuronextEntry(entries: IQuoteEntry[] | undefined): IQuoteEntry | undefined {
  const stocks = sblkStocks(entries);
  return stocks.find(e => EURONEXT_LIKE.test(label(e)))
    || stocks.find(e => !NASDAQ_LIKE.test(label(e)));
}
