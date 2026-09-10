import * as React from 'react';
import styles from './EventsWidget.module.scss';
import { events } from './data/mockData';

export interface IEventsWidgetState {
  index: number;
}

export default class EventsWidget extends React.Component<Record<string, never>, IEventsWidgetState> {
  constructor(props: Record<string, never>) {
    super(props);
    this.state = { index: 0 };
  }

  private _prev = (): void => {
    this.setState(prev => ({ index: (prev.index - 1 + events.length) % events.length }));
  };

  private _next = (): void => {
    this.setState(prev => ({ index: (prev.index + 1) % events.length }));
  };

  public render(): React.ReactElement {
    const event = events[this.state.index];

    return (
      <div className={styles.wrap}>
        <div className={styles.titleCard}>
          <h3 className={styles.title}>Upcoming Events</h3>
        </div>

        <div className={styles.eventCard}>
          <p className={styles.eventDate}>{event.date}</p>
          <p className={styles.eventTitle}>{event.title}</p>
          <div className={styles.mediaRow}>
            <button className={styles.arrowButton} onClick={this._prev} aria-label="Previous event">←</button>
            <div className={styles.eventImage}>{event.initials}</div>
            <button className={styles.arrowButton} onClick={this._next} aria-label="Next event">→</button>
          </div>
          <div className={styles.eventFooter}>
            <button className={styles.pill}>Read more →</button>
          </div>
        </div>
      </div>
    );
  }
}
