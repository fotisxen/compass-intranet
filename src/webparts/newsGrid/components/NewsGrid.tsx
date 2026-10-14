import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './NewsGrid.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import { NewsCardView } from '../../../shared/components/NewsCard';
import { type INewsItem, loadPromotedNews, loadPageByPath, normalizePageUrl } from '../../../shared/components/newsApi';

// Cards on screen at once: 1 large + 7 small.
const SLOTS = 8;
// How many of the newest promoted pages are fetched to rotate through.
const MAX_NEWS = 30;
// How often the unpinned cards move on by themselves.
const AUTO_ROTATE_MS = 5000;

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
  private _wrapRef = React.createRef<HTMLDivElement>();
  private _rotateTimer: number | undefined;

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

  public componentWillUnmount(): void {
    this._stopAutoRotate();
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
    this.setState({ pinnedItems, rotating: rest, start: 0 }, this._startAutoRotate);
  }

  /** Slots left over once the pinned cards have taken theirs. */
  private _freeSlots(): number {
    return Math.max(0, SLOTS - this.state.pinnedItems.length);
  }

  private _canRotate(): boolean {
    const free = this._freeSlots();
    return free > 0 && this.state.rotating.length > free;
  }

  private _stopAutoRotate(): void {
    if (this._rotateTimer !== undefined) {
      window.clearInterval(this._rotateTimer);
      this._rotateTimer = undefined;
    }
  }

  private _startAutoRotate = (): void => {
    this._stopAutoRotate();
    const reduceMotion = typeof window.matchMedia === 'function' && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (!this._canRotate() || reduceMotion) {
      return;
    }
    this._rotateTimer = window.setInterval(() => {
      if (!document.hidden && !this._isBeingUsed()) {
        this._step();
      }
    }, AUTO_ROTATE_MS);
  };

  // Asked fresh at every tick: a card someone is pointing at, or has focused,
  // never changes under them.
  private _isBeingUsed(): boolean {
    const el = this._wrapRef.current;
    return !!el && (el.matches(':hover') || el.contains(document.activeElement));
  }

  // Moves the unpinned cards on by one full set, wrapping round; pinned stay.
  private _step(): void {
    const free = this._freeSlots();
    this.setState(prev => {
      const n = prev.rotating.length;
      return n > free && free > 0 ? { start: (prev.start + free) % n } : null;
    });
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
        <div className={styles.wrap} ref={this._wrapRef}>
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
