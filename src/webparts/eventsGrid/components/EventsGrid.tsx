import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './EventsGrid.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import {
  type IEventItem,
  downloadIcs,
  eventUrl,
  formatMonthHeading,
  formatShort,
  loadEvents
} from '../../../shared/events/eventsApi';

export interface IEventsGridProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  listRef: string;
  eventPageUrl: string;
  maxEvents: number;
  emptyText: string;
}

interface IMonthGroup {
  heading: string;
  events: IEventItem[];
}

/** The events (already sorted by start) split into one group per calendar month. */
export function groupByMonth(events: IEventItem[]): IMonthGroup[] {
  const groups: IMonthGroup[] = [];
  events.forEach(ev => {
    const heading = formatMonthHeading(ev.start);
    const last = groups[groups.length - 1];
    if (last && last.heading === heading) {
      last.events.push(ev);
    } else {
      groups.push({ heading, events: [ev] });
    }
  });
  return groups;
}

export default function EventsGrid(props: IEventsGridProps): React.ReactElement {
  const { spHttpClient, siteUrl, listRef, eventPageUrl, maxEvents, emptyText } = props;
  const [events, setEvents] = React.useState<IEventItem[] | undefined>();
  const [failed, setFailed] = React.useState(false);

  React.useEffect(() => {
    ensureInterFont();
    loadEvents(spHttpClient, siteUrl, listRef, { top: maxEvents })
      .then(setEvents)
      .catch(() => setFailed(true));
  }, [spHttpClient, siteUrl, listRef, maxEvents]);

  if (failed) {
    return <div className={styles.page}><div className={styles.empty}>The events could not be loaded.</div></div>;
  }
  if (!events) {
    return <div className={styles.page} />;
  }
  if (events.length === 0) {
    return <div className={styles.page}><div className={styles.empty}>{emptyText}</div></div>;
  }

  return (
    <div className={styles.page}>
      {groupByMonth(events).map(group => (
        <section className={styles.month} key={group.heading}>
          <h3 className={styles.monthHeading}>{group.heading}</h3>
          <div className={styles.grid}>
            {group.events.map(ev => {
              const url = eventUrl(eventPageUrl, ev.id);
              return (
                <article className={styles.card} key={ev.id}>
                  <a className={styles.cardLink} href={url}>
                    <div className={styles.photo}>
                      {ev.imageUrl && <img src={ev.imageUrl} alt="" loading="lazy" />}
                    </div>
                    <div className={styles.body}>
                      {ev.category && <p className={styles.category}>{ev.category}</p>}
                      <h4 className={styles.title}>{ev.title}</h4>
                      <p className={styles.when}>{formatShort(ev)}</p>
                      {ev.location && <p className={styles.place}>{ev.location}</p>}
                    </div>
                  </a>
                  <div className={styles.actions}>
                    <button type="button" className={styles.calendarButton} onClick={() => downloadIcs(ev, url)}>
                      Add to calendar +
                    </button>
                  </div>
                </article>
              );
            })}
          </div>
        </section>
      ))}
    </div>
  );
}
