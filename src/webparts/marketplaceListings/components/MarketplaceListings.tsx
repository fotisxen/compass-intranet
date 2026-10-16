import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import shared from '../../../shared/marketplace/Marketplace.module.scss';
import styles from './MarketplaceListings.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import SetupNotice from '../../../shared/marketplace/SetupNotice';
import { ListingCard, ListingModal } from '../../../shared/marketplace/ListingView';
import { ensureMarketplaceLists, loadListings, type IListing } from '../../../shared/marketplace/marketplaceApi';
import { ALL_CATEGORIES, getFilter, onFilterChange, type IMarketplaceFilter } from '../../../shared/marketplace/marketplaceBus';

/** Cards shown at first, and how many more each "Load more" adds. */
const FIRST_PAGE = 6;
const PAGE_STEP = 3;

export interface IMarketplaceListingsProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  canManageLists: boolean;
  expiryDays: number;
}

type Status = 'loading' | 'ready' | 'missing' | 'unreadable';

/** Every searchable field of a listing, lowercased, for the search box. */
function searchText(l: IListing): string {
  return [l.title, l.price, l.area, l.details, l.category, l.authorName].join(' ').toLowerCase();
}

export function filterListings(listings: IListing[], filter: IMarketplaceFilter): IListing[] {
  const terms = filter.search.toLowerCase().split(/\s+/).filter(t => !!t);
  return listings.filter(l => {
    if (filter.category !== ALL_CATEGORIES && l.category.toLowerCase() !== filter.category.toLowerCase()) {
      return false;
    }
    const text = searchText(l);
    return terms.every(t => text.indexOf(t) !== -1);
  });
}

export default function MarketplaceListings(props: IMarketplaceListingsProps): React.ReactElement {
  const { spHttpClient, siteUrl, canManageLists, expiryDays } = props;
  const [status, setStatus] = React.useState<Status>('loading');
  const [listings, setListings] = React.useState<IListing[]>([]);
  const [filter, setFilterState] = React.useState<IMarketplaceFilter>(getFilter());
  const [shown, setShown] = React.useState<number>(FIRST_PAGE);
  const [open, setOpen] = React.useState<IListing | undefined>();

  const load = React.useCallback(async (): Promise<void> => {
    const result = await loadListings(spHttpClient, siteUrl, expiryDays);
    if (typeof result === 'string') {
      setStatus(result);
      return;
    }
    setListings(result.data);
    setStatus('ready');
  }, [spHttpClient, siteUrl, expiryDays]);

  React.useEffect(() => {
    ensureInterFont();
    load().catch(() => setStatus('unreadable'));
  }, [load]);

  React.useEffect(() => onFilterChange(next => {
    setFilterState(next);
    setShown(FIRST_PAGE);
  }), []);

  const setup = async (): Promise<void> => {
    await ensureMarketplaceLists(spHttpClient, siteUrl);
    await load();
  };

  if (status === 'missing') {
    return <div className={shared.page}><SetupNotice canManageLists={canManageLists} onSetup={setup} /></div>;
  }
  if (status === 'unreadable') {
    return <div className={shared.page}><div className={styles.empty}>The Marketplace listings could not be loaded.</div></div>;
  }

  const matching = filterListings(listings, filter);
  const visible = matching.slice(0, shown);

  return (
    <div className={shared.page}>
      {status === 'ready' && matching.length === 0 && (
        <div className={styles.empty}>{listings.length === 0 ? 'No listings yet — be the first to post one.' : 'No listings match your search.'}</div>
      )}
      <div className={styles.grid}>
        {visible.map(l => <ListingCard key={l.id} listing={l} onOpen={setOpen} />)}
      </div>
      {matching.length > shown && (
        <div className={styles.more}>
          <button type="button" className={shared.pill} style={{ fontSize: '0.65em' }} onClick={() => setShown(n => n + PAGE_STEP)}>
            Load more listings
          </button>
        </div>
      )}
      {open && <ListingModal listing={open} onClose={() => setOpen(undefined)} />}
    </div>
  );
}
