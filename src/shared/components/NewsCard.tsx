import * as React from 'react';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './NewsCard.module.scss';

// Real internal field name for the promoted news pages' category column,
// confirmed from the "Internal Company Announcements" page's own News web
// part filter config (Choice field, internal name "NewsCategory", display
// name "News Category").
const NEWS_CATEGORY_FIELD_NAME = 'NewsCategory';

export interface INewsCardProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  /** 1-based position among promoted news items, most recent first. Defaults to 1 (latest). */
  position?: number;
  /** Matches an INewsCategory label from the Newsroom sidebar's selection. */
  categoryFilter?: string;
}

interface ISpNewsItem {
  title: string;
  date: string;
  url: string;
  imageUrl?: string;
}

export interface INewsCardState {
  item?: ISpNewsItem;
}

interface ISpListItem {
  Title: string;
  FileRef: string;
  Created: string;
  // Site Pages' Banner Image column stores a JSON blob (serverRelativeUrl,
  // dimensions, crop info, ...), not a plain URL string.
  BannerImageUrl?: string;
}

interface IBannerImageField {
  serverRelativeUrl?: string;
}

function parseBannerImageUrl(raw?: string): string | undefined {
  if (!raw) {
    return undefined;
  }
  try {
    const parsed: IBannerImageField = JSON.parse(raw);
    return parsed.serverRelativeUrl || undefined;
  } catch {
    return undefined;
  }
}

interface ISpListItemsResponse {
  value: ISpListItem[];
}

export default class NewsCard extends React.Component<INewsCardProps, INewsCardState> {
  constructor(props: INewsCardProps) {
    super(props);
    this.state = { item: undefined };
  }

  public componentDidMount(): void {
    this._loadItem().catch(() => {
      // Real news couldn't be loaded (no promoted pages yet, permissions,
      // network) — leave the card empty rather than show a fabricated item.
    });
  }

  public componentDidUpdate(prevProps: INewsCardProps): void {
    if (prevProps.position === this.props.position && prevProps.categoryFilter === this.props.categoryFilter) {
      return;
    }

    this.setState({ item: undefined });
    this._loadItem().catch(() => {
      // Same empty-on-failure behavior as componentDidMount.
    });
  }

  private async _loadItem(): Promise<void> {
    const { spHttpClient, siteUrl, categoryFilter } = this.props;
    const position = this.props.position || 1;

    // "Site Pages" filtered to PromotedState eq 2 is how SharePoint marks a
    // page as posted to Newsroom/News. Fetching the top `position` items
    // (most recent first) and taking the last one gives the Nth most
    // recent promoted article, so dropping this web part multiple times
    // with position 1, 2, 3... reproduces the original 3-card grid.
    let filter = 'PromotedState eq 2';
    if (categoryFilter) {
      filter += ` and ${NEWS_CATEGORY_FIELD_NAME} eq '${categoryFilter}'`;
    }

    const endpoint =
      `${siteUrl}/_api/web/lists/GetByTitle('Site Pages')/items` +
      `?$select=Title,FileRef,Created,BannerImageUrl&$filter=${filter}&$orderby=Created desc&$top=${position}`;

    const response: SPHttpClientResponse = await spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

    if (!response.ok) {
      return;
    }

    const data: ISpListItemsResponse = await response.json();

    if (!data.value || data.value.length < position) {
      return;
    }

    const spItem = data.value[position - 1];
    this.setState({
      item: {
        title: spItem.Title,
        date: new Date(spItem.Created).toLocaleDateString('en-US', { day: 'numeric', month: 'long', year: 'numeric' }),
        url: spItem.FileRef,
        imageUrl: parseBannerImageUrl(spItem.BannerImageUrl)
      }
    });
  }

  public render(): React.ReactElement {
    const { item } = this.state;

    if (!item) {
      return <></>;
    }

    const thumbStyle = item.imageUrl ? { backgroundImage: `url(${item.imageUrl})` } : undefined;

    return (
      <div className={styles.card}>
        <div className={styles.thumb} style={thumbStyle} />
        <div className={styles.body}>
          <h4 className={styles.title}>{item.title}</h4>
          <div className={styles.footerRow}>
            <span className={styles.date}>{item.date}</span>
            <a className={styles.pill} href={item.url}>Read more →</a>
          </div>
        </div>
      </div>
    );
  }
}
