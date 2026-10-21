import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './EventImage.module.scss';
import { useEvent } from '../../../shared/events/useEvent';

export interface IEventImageProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  listRef: string;
  aspectRatio: string;
  radius: number;
  isEditMode: boolean;
}

export default function EventImage(props: IEventImageProps): React.ReactElement {
  const state = useEvent(props.spHttpClient, props.siteUrl, props.listRef, props.isEditMode);
  if (state.status === 'none') {
    // No event to show: while editing keep a visible block so the web part can be found.
    return props.isEditMode
      ? <div className={styles.page}><div className={styles.frame} style={{ aspectRatio: props.aspectRatio, borderRadius: props.radius }}><span className={styles.hint}>Event image — no event found</span></div></div>
      : <div className={styles.page} />;
  }

  const event = state.status === 'ready' ? state.event : undefined;
  return (
    <div className={styles.page}>
      <div className={styles.frame} style={{ aspectRatio: props.aspectRatio, borderRadius: props.radius }}>
        {event && event.imageUrl && <img src={event.imageUrl} alt={event.title} />}
        {props.isEditMode && state.status === 'ready' && state.preview && <span className={styles.hint}>Event image — showing the next event; the real page uses ?id=</span>}
      </div>
    </div>
  );
}
