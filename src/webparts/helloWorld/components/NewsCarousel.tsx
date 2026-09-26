import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './NewsCarousel.module.scss';
import { NewsCardView } from '../../../shared/components/NewsCard';
import { type INewsItem, loadPromotedNews, loadPageByPath, normalizePageUrl } from '../../../shared/components/newsApi';

// How many of the newest promoted pages the carousel can page through.
const MAX_NEWS = 30;
// Cards visible at once (first one wide, see .grid in the scss).
const VISIBLE = 3;

export interface INewsCarouselProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  /** Pages pinned for everyone, shown first in this order — full URLs, server-relative paths, or just a Site Pages file name. Set in the web part's property pane. */
  pinned: string[];
}

export interface INewsCarouselState {
  items: INewsItem[];
  start: number;
  /** urls of the items that came from the pinned list. */
  pinnedUrls: string[];
}

export default class NewsCarousel extends React.Component<INewsCarouselProps, INewsCarouselState> {
  constructor(props: INewsCarouselProps) {
    super(props);
    this.state = { items: [], start: 0, pinnedUrls: [] };
  }

  public componentDidMount(): void {
    this._load().catch(() => {
      // News couldn't be loaded — leave the row empty rather than show
      // made-up articles.
    });
  }

  public componentDidUpdate(prev: INewsCarouselProps): void {
    if (prev.pinned.join('|') !== this.props.pinned.join('|')) {
      this._load().catch(() => {
        // Same as componentDidMount.
      });
    }
  }

  private async _load(): Promise<void> {
    const { spHttpClient, siteUrl, pinned } = this.props;
    const promoted = await loadPromotedNews(spHttpClient, siteUrl, MAX_NEWS);
    const byPath = new Map<string, INewsItem>();
    promoted.forEach(n => byPath.set(normalizePageUrl(siteUrl, n.url), n));

    const pinnedItems: INewsItem[] = [];
    const seen = new Set<string>();
    for (const raw of pinned) {
      const path = normalizePageUrl(siteUrl, raw);
      if (!path || seen.has(path)) {
        continue;
      }
      seen.add(path);
      // Pinned pages don't have to be promoted news, so fetch the page
      // itself when it isn't among the promoted ones.
      const item = byPath.get(path) || (await loadPageByPath(spHttpClient, siteUrl, path).catch(() => undefined));
      if (item) {
        pinnedItems.push(item);
      }
    }

    const rest = promoted.filter(n => !seen.has(normalizePageUrl(siteUrl, n.url)));
    this.setState({ items: [...pinnedItems, ...rest], start: 0, pinnedUrls: pinnedItems.map(n => n.url) });
  }

  private _prev = (): void => {
    this.setState(prev => ({ start: Math.max(0, prev.start - 1) }));
  };

  private _next = (): void => {
    this.setState(prev => ({ start: Math.min(Math.max(0, prev.items.length - VISIBLE), prev.start + 1) }));
  };

  public render(): React.ReactElement {
    const { items, start, pinnedUrls } = this.state;
    const visible = items.slice(start, start + VISIBLE);
    const canScroll = items.length > VISIBLE;

    return (
      <div className={styles.wrap}>
        <div className={styles.grid}>
          {visible.map(item => (
            <NewsCardView key={item.url} item={item} pinned={pinnedUrls.indexOf(item.url) >= 0} />
          ))}
        </div>

        {canScroll && (
          <>
            <button
              className={`${styles.arrow} ${styles.arrowLeft}`}
              onClick={this._prev}
              disabled={start === 0}
              aria-label="Previous news"
            >
              ←
            </button>
            <button
              className={`${styles.arrow} ${styles.arrowRight}`}
              onClick={this._next}
              disabled={start >= items.length - VISIBLE}
              aria-label="Next news"
            >
              →
            </button>
          </>
        )}
      </div>
    );
  }
}
