import * as React from 'react';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './NewsCard.module.scss';
import { news as mockNews } from './data/newsMockData';

export interface INewsCardProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  /** 1-based position among promoted news items, most recent first. Defaults to 1 (latest). */
  position?: number;
}

interface ISpNewsItem {
  title: string;
  date: string;
  url: string;
  imageUrl?: string;
}

export interface INewsCardState {
  item: ISpNewsItem;
}

interface ISpListItem {
  Title: string;
  FileRef: string;
  Created: string;
}

interface ISpListItemsResponse {
  value: ISpListItem[];
}

function mockItemAt(position: number): ISpNewsItem {
  const item = mockNews[(position - 1) % mockNews.length];
  return { ...item, url: '#' };
}

export default class NewsCard extends React.Component<INewsCardProps, INewsCardState> {
  constructor(props: INewsCardProps) {
    super(props);
    this.state = { item: mockItemAt(props.position || 1) };
  }

  public componentDidMount(): void {
    this._loadItem().catch(() => {
      // Real news couldn't be loaded (no promoted pages yet, permissions,
      // network) — the mock item already set in state stays as a fallback
      // so the card never renders empty.
    });
  }

  public componentDidUpdate(prevProps: INewsCardProps): void {
    if (prevProps.position === this.props.position) {
      return;
    }

    this.setState({ item: mockItemAt(this.props.position || 1) });
    this._loadItem().catch(() => {
      // Same silent-fallback as componentDidMount.
    });
  }

  private async _loadItem(): Promise<void> {
    const { spHttpClient, siteUrl } = this.props;
    const position = this.props.position || 1;

    // "Site Pages" filtered to PromotedState eq 2 is how SharePoint marks a
    // page as posted to Newsroom/News. Fetching the top `position` items
    // (most recent first) and taking the last one gives the Nth most
    // recent promoted article, so dropping this web part multiple times
    // with position 1, 2, 3... reproduces the original 3-card grid.
    const endpoint =
      `${siteUrl}/_api/web/lists/GetByTitle('Site Pages')/items` +
      `?$select=Title,FileRef,Created&$filter=PromotedState eq 2&$orderby=Created desc&$top=${position}`;

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
        url: spItem.FileRef
      }
    });
  }

  public render(): React.ReactElement {
    const { item } = this.state;

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
