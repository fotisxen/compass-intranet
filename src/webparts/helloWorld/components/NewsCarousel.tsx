import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './NewsCarousel.module.scss';
import ArrowIcon from '../../../shared/components/ArrowIcon';
import { NewsCardView } from '../../../shared/components/NewsCard';
import { type INewsItem, loadPromotedNews, loadPageByPath, normalizePageUrl } from '../../../shared/components/newsApi';

// How many of the newest promoted pages the carousel can page through.
const MAX_NEWS = 30;
// Cards visible at once (first one wide, see .grid in the scss): 3 on desktop
// and tablet, 2 on a phone (the same width as the scss's phone layout).
const VISIBLE = 3;
const PHONE_VISIBLE = 2;
const PHONE_QUERY = '(max-width: 600px)';

function visibleSlots(): number {
  return typeof window.matchMedia === 'function' && window.matchMedia(PHONE_QUERY).matches ? PHONE_VISIBLE : VISIBLE;
}
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
  /** How many cards fit on screen at the current width. */
  slots: number;
}

export default class NewsCarousel extends React.Component<INewsCarouselProps, INewsCarouselState> {
  constructor(props: INewsCarouselProps) {
    super(props);
    this.state = { pinnedItems: [], rotating: [], start: 0, slots: visibleSlots() };
  }

  private _wrapRef = React.createRef<HTMLDivElement>();
  private _rotateTimer: number | undefined;

  public componentDidMount(): void {
    window.addEventListener('resize', this._onResize);
    this._load().catch(() => {
      // News couldn't be loaded — leave the row empty rather than show
      // made-up articles.
    });
  }

  public componentWillUnmount(): void {
    window.removeEventListener('resize', this._onResize);
    this._stopAutoRotate();
  }

  // Turning a tablet sideways or resizing the window can change how many cards
  // fit; start again from the first set so no card is skipped.
  private _onResize = (): void => {
    const slots = visibleSlots();
    if (slots !== this.state.slots) {
      this.setState({ slots, start: 0 }, this._startAutoRotate);
    }
  };

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

  /**
   * What stays put and what rotates. On desktop and tablet the pinned cards
   * keep their slots and only the others rotate. On a phone there are only 2
   * slots, so pinned cards would leave nothing (or one card) to rotate and
   * the rest of the news could never be reached — there they simply come
   * first and then page through together with everything else.
   */
  private _lists(): { fixed: INewsItem[]; rotating: INewsItem[] } {
    const { pinnedItems, rotating, slots } = this.state;
    if (slots === PHONE_VISIBLE) {
      return { fixed: [], rotating: [...pinnedItems, ...rotating] };
    }
    return { fixed: pinnedItems.slice(0, slots), rotating };
  }

  /** Slots left over once the pinned cards have taken theirs. */
  private _freeSlots(): number {
    return Math.max(0, this.state.slots - this._lists().fixed.length);
  }

  private _canRotate(): boolean {
    return this._lists().rotating.length > this._freeSlots() && this._freeSlots() > 0;
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
    const n = this._lists().rotating.length;
    this.setState(prev => (n > free && free > 0 ? { start: (((prev.start + direction * free) % n) + n) % n } : null));
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
    const { start } = this.state;
    const { fixed, rotating } = this._lists();
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
