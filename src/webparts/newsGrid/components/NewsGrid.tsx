import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './NewsGrid.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import { NewsCardView } from '../../../shared/components/NewsCard';
import { type INewsItem, loadPromotedNews, loadPageByPath, normalizePageUrl } from '../../../shared/components/newsApi';

// Cards on screen at once: 3 rows of 3, all the same size.
const SLOTS = 9;
// The newest promoted pages fetched; the grid always shows the newest ones that fit.
const MAX_NEWS = SLOTS + 10;

export interface INewsGridProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  /** Pages pinned for everyone — full URLs, server-relative paths, or just a Site Pages file name. They fill the first slots and never move. */
  pinnedNews: string[];
}

export interface INewsGridState {
  pinnedItems: INewsItem[];
  /** Everything else, newest first. */
  rotating: INewsItem[];
  /** Index into "rotating" of the first rotating card on screen. */
  start: number;
}

export default class NewsGrid extends React.Component<INewsGridProps, INewsGridState> {

  constructor(props: INewsGridProps) {
    super(props);
    this.state = { pinnedItems: [], rotating: [], start: 0 };
  }

  public componentDidMount(): void {
    ensureInterFont();
    this._load().catch(() => {
      // News couldn't be loaded — leave the grid empty rather than show made-up articles.
    });
  }

  public componentDidUpdate(prev: INewsGridProps): void {
    if (prev.pinnedNews.join('|') !== this.props.pinnedNews.join('|')) {
      this._load().catch(() => {
        // Same as componentDidMount.
      });
    }
  }

  private async _load(): Promise<void> {
    const { spHttpClient, siteUrl, pinnedNews } = this.props;
    const promoted = await loadPromotedNews(spHttpClient, siteUrl, MAX_NEWS);
    const byPath = new Map<string, INewsItem>();
    promoted.forEach(n => byPath.set(normalizePageUrl(siteUrl, n.url), n));

    const pinnedItems: INewsItem[] = [];
    const seen = new Set<string>();
    for (const raw of pinnedNews) {
      const path = normalizePageUrl(siteUrl, raw);
      if (!path || seen.has(path)) {
        continue;
      }
      seen.add(path);
      // Pinned pages don't have to be promoted news, so fetch the page itself
      // when it isn't among the promoted ones.
      const item = byPath.get(path) || (await loadPageByPath(spHttpClient, siteUrl, path).catch(() => undefined));
      if (item) {
        pinnedItems.push(item);
      }
    }

    const rest = promoted.filter(n => !seen.has(normalizePageUrl(siteUrl, n.url)));
    this.setState({ pinnedItems, rotating: rest, start: 0 });
  }

  /** Slots left over once the pinned cards have taken theirs. */
  private _freeSlots(): number {
    return Math.max(0, SLOTS - this.state.pinnedItems.length);
  }

  public render(): React.ReactElement {
    const { pinnedItems, rotating, start } = this.state;
    const fixed = pinnedItems.slice(0, SLOTS);
    const free = this._freeSlots();
    const shown: INewsItem[] = [];
    for (let i = 0; i < Math.min(free, rotating.length); i++) {
      shown.push(rotating[(start + i) % rotating.length]);
    }

    return (
      <div className={styles.page}>
        <div className={styles.wrap}>
          <div className={styles.grid}>
            {fixed.map(item => (
              <NewsCardView key={item.url} item={item} pinned={true} />
            ))}
            {shown.map(item => (
              <NewsCardView key={item.url} item={item} />
            ))}
          </div>
        </div>
      </div>
    );
  }
}
