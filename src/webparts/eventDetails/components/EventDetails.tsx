import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './EventDetails.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import { formatWhen, sanitizeHtml } from '../../../shared/events/eventsApi';
import { useEvent } from '../../../shared/events/useEvent';

export interface IEventDetailsProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  listRef: string;
  showTitle: boolean;
  /** Pixels at the 1440px reference width (scales with the page). */
  titleSize: number;
  showWhen: boolean;
  showTimezone: boolean;
  showCategory: boolean;
  showLocation: boolean;
  showAbout: boolean;
  /** Space kept free on the left, in pixels. */
  spaceLeft: number;
  isEditMode: boolean;
}

export default function EventDetails(props: IEventDetailsProps): React.ReactElement {
  const state = useEvent(props.spHttpClient, props.siteUrl, props.listRef, props.isEditMode);

  React.useEffect(() => {
    ensureInterFont();
  }, []);

  if (state.status === 'none') {
    return (
      <div className={styles.page} style={{ paddingLeft: props.spaceLeft }}>
        {props.isEditMode && <span className={styles.hint}>Event details — no event found. Open the page with ?eventId=&lt;event id&gt; at the end of its address.</span>}
      </div>
    );
  }
  if (state.status === 'loading') {
    return <div className={styles.page} />;
  }

  const { event, preview } = state;
  const about = props.showAbout ? sanitizeHtml(event.descriptionHtml).trim() : '';
  const whenLines = formatWhen(event, props.showTimezone);

  return (
    <div className={styles.page} style={{ paddingLeft: props.spaceLeft }}>
      {props.isEditMode && preview && <span className={styles.hint}>Showing the next upcoming event — on the real page the event comes from ?eventId= in the address.</span>}
      {props.showTitle && <h1 className={styles.title} style={{ fontSize: `${props.titleSize / 20}em` }}>{event.title}</h1>}
      <div className={styles.rows}>
        {props.showWhen && (
          <>
            <div className={styles.label}>When</div>
            <div className={styles.value}>
              {whenLines.map((line, i) => (
                <span key={i} className={`${styles.line} ${i === whenLines.length - 1 && props.showTimezone && whenLines.length > 1 ? styles.lineGap : ''}`}>{line}</span>
              ))}
            </div>
          </>
        )}
        {props.showCategory && event.category && (
          <>
            <div className={styles.label}>Category</div>
            <div className={styles.value}>{event.category}</div>
          </>
        )}
        {props.showLocation && event.location && (
          <>
            <div className={styles.label}>Location</div>
            <div className={styles.value}>{event.location}</div>
          </>
        )}
        {about && (
          <>
            <div className={styles.label}>About</div>
            <div className={`${styles.value} ${styles.about}`} dangerouslySetInnerHTML={{ __html: about }} />
          </>
        )}
      </div>
    </div>
  );
}
