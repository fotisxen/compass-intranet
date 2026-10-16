import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import shared from '../../../shared/marketplace/Marketplace.module.scss';
import styles from './MarketplaceFilters.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import SetupNotice from '../../../shared/marketplace/SetupNotice';
import { ensureMarketplaceLists, loadCategories, type ICategory } from '../../../shared/marketplace/marketplaceApi';
import { ALL_CATEGORIES, getFilter, setFilter } from '../../../shared/marketplace/marketplaceBus';

export interface IMarketplaceFiltersProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  canManageLists: boolean;
}

type Status = 'loading' | 'ready' | 'missing' | 'unreadable';

export default function MarketplaceFilters(props: IMarketplaceFiltersProps): React.ReactElement {
  const { spHttpClient, siteUrl, canManageLists } = props;
  const [status, setStatus] = React.useState<Status>('loading');
  const [categories, setCategories] = React.useState<ICategory[]>([]);
  const [category, setCategory] = React.useState<string>(getFilter().category);
  const [search, setSearch] = React.useState<string>(getFilter().search);

  const load = React.useCallback(async (): Promise<void> => {
    const result = await loadCategories(spHttpClient, siteUrl);
    if (typeof result === 'string') {
      setStatus(result);
      return;
    }
    setCategories(result.data);
    setStatus('ready');
  }, [spHttpClient, siteUrl]);

  React.useEffect(() => {
    ensureInterFont();
    load().catch(() => setStatus('unreadable'));
  }, [load]);

  const choose = (name: string): void => {
    setCategory(name);
    setFilter({ category: name });
  };

  const onSearch = (e: React.ChangeEvent<HTMLInputElement>): void => {
    setSearch(e.target.value);
    setFilter({ search: e.target.value });
  };

  const setup = async (): Promise<void> => {
    await ensureMarketplaceLists(spHttpClient, siteUrl);
    await load();
  };

  if (status === 'missing') {
    return <div className={shared.page}><SetupNotice canManageLists={canManageLists} onSetup={setup} /></div>;
  }

  return (
    <div className={shared.page}>
      <div className={styles.chips}>
        {[{ id: 0, name: ALL_CATEGORIES }, ...categories].map(c => (
          <button
            key={c.id}
            type="button"
            className={`${styles.chip} ${category === c.name ? styles.chipActive : ''}`}
            onClick={() => choose(c.name)}
          >
            {c.name}
          </button>
        ))}
      </div>
      <div className={styles.search}>
        <input
          className={styles.searchInput}
          type="text"
          value={search}
          onChange={onSearch}
          placeholder="Search marketplace..."
          aria-label="Search the marketplace"
        />
        <svg className={styles.searchIcon} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" aria-hidden="true">
          <circle cx="10.5" cy="10.5" r="6.5" />
          <path d="M16 16l5 5" />
        </svg>
      </div>
    </div>
  );
}
