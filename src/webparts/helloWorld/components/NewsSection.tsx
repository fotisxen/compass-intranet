import * as React from 'react';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './NewsSection.module.scss';
import { news as mockNews, newsCategories, type INewsItem } from './data/mockData';

export interface INewsSectionProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
}

interface ISpNewsItem extends INewsItem {
  url: string;
}

export interface INewsSectionState {
  items: ISpNewsItem[];
  isUsingMockData: boolean;
}

interface ISpListItem {
  Title: string;
  FileRef: string;
  Created: string;
}

interface ISpListItemsResponse {
  value: ISpListItem[];
}

const PAGE_SIZE = 3;

export default class NewsSection extends React.Component<INewsSectionProps, INewsSectionState> {
  constructor(props: INewsSectionProps) {
    super(props);
    this.state = {
      items: mockNews.map(n => ({ ...n, url: '#' })),
      isUsingMockData: true
    };
  }

  public componentDidMount(): void {
    this._loadNews().catch(() => {
      // Real news couldn't be loaded (no promoted pages yet, permissions,
      // network) — the mock data already set in state stays as a fallback
      // so the section never renders empty.
    });
  }

  private async _loadNews(): Promise<void> {
    const { spHttpClient, siteUrl } = this.props;

    // "Site Pages" filtered to PromotedState eq 2 is how SharePoint marks a
    // page as posted to Newsroom/News — this is the same REST endpoint the
    // out-of-the-box News web part reads from.
    const endpoint =
      `${siteUrl}/_api/web/lists/GetByTitle('Site Pages')/items` +
      `?$select=Title,FileRef,Created&$filter=PromotedState eq 2&$orderby=Created desc&$top=${PAGE_SIZE}`;

    const response: SPHttpClientResponse = await spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

    if (!response.ok) {
      return;
    }

    const data: ISpListItemsResponse = await response.json();

    if (!data.value || data.value.length === 0) {
      return;
    }

    const items: ISpNewsItem[] = data.value.map(item => ({
      title: item.Title,
      date: new Date(item.Created).toLocaleDateString('en-US', { day: 'numeric', month: 'long', year: 'numeric' }),
      url: item.FileRef
    }));

    this.setState({ items, isUsingMockData: false });
  }

  public render(): React.ReactElement {
    const { items } = this.state;

    return (
      <div className={styles.wrap}>
        <div className={styles.sidebar}>
          <h3 className={styles.sidebarTitle}>Newsroom</h3>
          <div className={styles.tabs}>
            {newsCategories.map(cat => (
              <a
                key={cat.label}
                href="#"
                onClick={e => e.preventDefault()}
                className={
                  cat.highlighted
                    ? styles.tabHighlighted
                    : `${styles.tab} ${cat.label === 'Corporate' ? styles.tabActive : ''}`
                }
              >
                {cat.label}
              </a>
            ))}
          </div>
        </div>

        <div className={styles.grid}>
          {items.map(n => (
            <div className={styles.card} key={n.title}>
              <div className={styles.thumb} />
              <div className={styles.body}>
                <h4 className={styles.title}>{n.title}</h4>
                <div className={styles.footerRow}>
                  <span className={styles.date}>{n.date}</span>
                  <a className={styles.pill} href={n.url}>Read more →</a>
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }
}
