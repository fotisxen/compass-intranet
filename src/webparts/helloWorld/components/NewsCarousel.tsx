import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './NewsCarousel.module.scss';
import ArrowIcon from '../../../shared/components/ArrowIcon';
import { NewsCardView } from '../../../shared/components/NewsCard';
import { type INewsItem, loadPromotedNews, loadPageByPath, normalizePageUrl } from '../../../shared/components/newsApi';

// How many of the newest promoted pages the carousel can page through.
const MAX_NEWS = 30;
// Cards visible at once (first one wide, see .grid in the scss).
const VISIBLE = 3;
// How often the unpinned cards move on by themselves.
const AUTO_ROTATE_MS = 5000;

export interface INewsCarouselProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  /** Pages pinned for everyone, shown first in this order — full URLs, server-relative paths, or just a Site Pages file name. Set in the web part's property pane. */
  pinned: string[];
}

export interface INewsCarouselState {
  /** Pinned pages, in the order set in the property pane — they never move. */
  pinnedItems: INewsItem[];
  /** Everything else (newest first) — cycles through the slots the pinned cards leave free. */
  rotating: INewsItem[];
  /** Index into "rotating" of the first rotating card on screen. */
  start: number;
}

export default class NewsCarousel extends React.Component<INewsCarouselProps, INewsCarouselState> {
  constructor(props: INewsCarouselProps) {
    super(props);
    this.state = { pinnedItems: [], rotating: [], start: 0 };
  }

  private _wrapRef = React.createRef<HTMLDivElement>();
  private _rotateTimer: number | undefined;

  public componentDidMount(): void {
    this._load().catch(() => {
      // News couldn't be loaded — leave the row empty rather than show
      // made-up articles.
    });
  }

  public componentWillUnmount(): void {
    this._stopAutoRotate();
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
    this.setState({ pinnedItems, rotating: rest, start: 0 }, this._startAutoRotate);
  }

  /** Slots left over once the pinned cards have taken theirs. */
  private _freeSlots(): number {
    return Math.max(0, VISIBLE - this.state.pinnedItems.length);
  }

  private _canRotate(): boolean {
    return this.state.rotating.length > this._freeSlots() && this._freeSlots() > 0;
  }

  private _stopAutoRotate(): void {
    if (this._rotateTimer !== undefined) {
      window.clearInterval(this._rotateTimer);
      this._rotateTimer = undefined;
    }
  }

  // (Re)starts the countdown — after load and after every manual click, so a
  // click is never followed by an instant jump.
  private _startAutoRotate = (): void => {
    this._stopAutoRotate();
    const reduceMotion = typeof window.matchMedia === 'function' && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (!this._canRotate() || reduceMotion) {
      return;
    }
    this._rotateTimer = window.setInterval(() => {
      if (!document.hidden && !this._isBeingUsed()) {
        this._step(1);
      }
    }, AUTO_ROTATE_MS);
  };

  // Asked fresh at every tick: the card someone is pointing at, or has
  // focused, never changes under them.
  private _isBeingUsed(): boolean {
    const el = this._wrapRef.current;
    return !!el && (el.matches(':hover') || el.contains(document.activeElement));
  }

  // Moves the unpinned cards one full set forward/back (wrapping round),
  // leaving the pinned ones where they are.
  private _step(direction: 1 | -1): void {
    const free = this._freeSlots();
    this.setState(prev => {
      const n = prev.rotating.length;
      return n > free && free > 0 ? { start: (((prev.start + direction * free) % n) + n) % n } : null;
    });
  }

  private _prev = (): void => {
    this._step(-1);
    this._startAutoRotate();
  };

  private _next = (): void => {
    this._step(1);
    this._startAutoRotate();
  };

  public render(): React.ReactElement {
    const { pinnedItems, rotating, start } = this.state;
    const fixed = pinnedItems.slice(0, VISIBLE);
    const free = this._freeSlots();
    const shown: INewsItem[] = [];
    for (let i = 0; i < Math.min(free, rotating.length); i++) {
      shown.push(rotating[(start + i) % rotating.length]);
    }
    const canScroll = this._canRotate();

    return (
      <div className={styles.wrap} ref={this._wrapRef}>
        <div className={styles.grid}>
          {fixed.map(item => (
            <NewsCardView key={item.url} item={item} pinned={true} />
          ))}
          {shown.map(item => (
            <NewsCardView key={item.url} item={item} />
          ))}
        </div>

        {canScroll && (
          <>
            <button
              className={`${styles.arrow} ${styles.arrowLeft}`}
              onClick={this._prev}
              aria-label="Previous news"
            >
              <ArrowIcon direction="left" />
            </button>
            <button
              className={`${styles.arrow} ${styles.arrowRight}`}
              onClick={this._next}
              aria-label="Next news"
            >
              <ArrowIcon direction="right" />
            </button>
          </>
        )}
      </div>
    );
  }
}
