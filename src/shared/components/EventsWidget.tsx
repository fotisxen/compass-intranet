import * as React from 'react';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './EventsWidget.module.scss';
import { type IEvent } from './data/eventsMockData';

export interface IEventsWidgetProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
}

export interface IEventsWidgetState {
  index: number;
  events: IEvent[];
}

// The real "Upcoming Events" widget on the live site reads from a plain
// SharePoint Events list (Title + EventDate) — this is that list's GUID,
// found in the page's own web part configuration.
const UPCOMING_EVENTS_LIST_ID = 'dd3316a0-3bc0-4d4c-acff-71851299dad7';
const PAGE_SIZE = 10;

interface IThumbnailField {
  Description?: string;
  Url?: string;
  serverRelativeUrl?: string;
}

interface ISpEventItem {
  Id: number;
  Title: string;
  EventDate: string;
  // "Thumbnail"-type column. Classic REST has been observed returning this
  // both as a JSON-encoded blob (serverRelativeUrl, crop info, ...) and as
  // an already-parsed { Description, Url } object — handle both shapes.
  BannerUrl?: string | IThumbnailField;
}

interface ISpListItemsResponse {
  value: ISpEventItem[];
}

function parseThumbnailUrl(raw?: string | IThumbnailField): string | undefined {
  if (!raw) {
    return undefined;
  }
  if (typeof raw === 'object') {
    return raw.Url || raw.serverRelativeUrl || raw.Description || undefined;
  }
  const trimmed = raw.trim();
  if (!trimmed) {
    return undefined;
  }
  if (trimmed.charAt(0) === '{') {
    try {
      const parsed: IThumbnailField = JSON.parse(trimmed);
      return parsed.Url || parsed.serverRelativeUrl || parsed.Description || undefined;
    } catch {
      return undefined;
    }
  }
  return trimmed;
}

function pad2(n: number): string {
  return n < 10 ? `0${n}` : `${n}`;
}

function formatDisplayDate(iso: string): string {
  const d = new Date(iso);
  return `${pad2(d.getDate())}.${pad2(d.getMonth() + 1)}.${d.getFullYear()}`;
}

function initialsFor(title: string): string {
  const initials = title
    .trim()
    .split(/\s+/)
    .slice(0, 2)
    .map(w => w.charAt(0).toUpperCase())
    .join('');
  return initials || 'EV';
}

export default class EventsWidget extends React.Component<IEventsWidgetProps, IEventsWidgetState> {
  constructor(props: IEventsWidgetProps) {
    super(props);
    this.state = { index: 0, events: [] };
  }

  public componentDidMount(): void {
    this._loadEvents().catch(() => {
      // Real list couldn't be loaded — leave events empty rather than
      // show fabricated ones.
    });
  }

  private async _loadEvents(): Promise<void> {
    const { spHttpClient, siteUrl } = this.props;
    const today = new Date().toISOString();
    const endpoint =
      `${siteUrl}/_api/web/lists(guid'${UPCOMING_EVENTS_LIST_ID}')/items` +
      `?$select=Id,Title,EventDate,BannerUrl&$filter=EventDate ge datetime'${today}'&$orderby=EventDate asc&$top=${PAGE_SIZE}`;

    const response: SPHttpClientResponse = await spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    if (!response.ok) {
      return;
    }

    const data: ISpListItemsResponse = await response.json();
    if (!data.value || data.value.length === 0) {
      return;
    }

    const events: IEvent[] = data.value.map(item => ({
      date: formatDisplayDate(item.EventDate),
      title: item.Title,
      initials: initialsFor(item.Title),
      imageUrl: parseThumbnailUrl(item.BannerUrl),
      // Calendar-list items open through Event.aspx (not the generic
      // listform.aspx display form) — confirmed against the real event's
      // own "Read more" URL on the live site.
      url: `${siteUrl}/_layouts/15/Event.aspx?ListGuid=${UPCOMING_EVENTS_LIST_ID}&ItemId=${item.Id}`
    }));

    this.setState({ events, index: 0 });
  }

  private _prev = (): void => {
    if (this.state.events.length === 0) {
      return;
    }
    this.setState(prev => ({ index: (prev.index - 1 + prev.events.length) % prev.events.length }));
  };

  private _next = (): void => {
    if (this.state.events.length === 0) {
      return;
    }
    this.setState(prev => ({ index: (prev.index + 1) % prev.events.length }));
  };

  public render(): React.ReactElement {
    const { events, index } = this.state;
    const event = events[index];
    // With only one (or zero) upcoming events, cycling always lands back on
    // the same index — the arrows would look broken rather than just idle,
    // so hide them instead of showing dead controls.
    const hasMultiple = events.length > 1;

    return (
      <div className={styles.wrap}>
        <div className={styles.titleCard}>
          <h3 className={styles.title}>Upcoming Events</h3>
        </div>

        {event ? (
          <div className={styles.eventCard}>
            <div className={styles.textGroup}>
              <p className={styles.eventDate}>{event.date}</p>
              <p className={styles.eventTitle}>{event.title}</p>
            </div>
            <div className={styles.mediaRow} style={hasMultiple ? undefined : { justifyContent: 'center' }}>
              {hasMultiple && (
                <button className={styles.arrowButton} onClick={this._prev} aria-label="Previous event">←</button>
              )}
              <div className={styles.eventImage} style={event.imageUrl ? { backgroundImage: `url("${event.imageUrl}")` } : undefined}>
                {!event.imageUrl && event.initials}
              </div>
              {hasMultiple && (
                <button className={styles.arrowButton} onClick={this._next} aria-label="Next event">→</button>
              )}
            </div>
            <div className={styles.eventFooter}>
              <a className={styles.pill} href={event.url}>Read more →</a>
            </div>
          </div>
        ) : (
          <div className={styles.eventCard}>
            <p className={styles.eventDate}>No upcoming events</p>
          </div>
        )}
      </div>
    );
  }
}
