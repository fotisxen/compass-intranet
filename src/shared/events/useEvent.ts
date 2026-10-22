import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import { type IEventItem, eventIdFromUrl, loadEvent, loadNextEvent } from './eventsApi';

export type EventState =
  | { status: 'loading' }
  | { status: 'none' }
  | { status: 'ready'; event: IEventItem; preview: boolean };

/**
 * The event a detail page was opened for (?eventId=…). While editing a page that has
 * no ?eventId= yet it shows the next upcoming event instead, so the author can see
 * the web part; a visitor without an id simply sees nothing.
 */
export function useEvent(client: SPHttpClient, siteUrl: string, listRef: string, isEditMode: boolean): EventState {
  const [state, setState] = React.useState<EventState>({ status: 'loading' });

  React.useEffect(() => {
    let cancelled = false;
    const id = eventIdFromUrl();
    const load = id !== undefined
      ? loadEvent(client, siteUrl, listRef, id).then(event => (event ? { event, preview: false } : undefined))
      : isEditMode
        ? loadNextEvent(client, siteUrl, listRef).then(event => (event ? { event, preview: true } : undefined))
        : Promise.resolve(undefined);
    load
      .then(found => {
        if (!cancelled) {
          setState(found ? { status: 'ready', event: found.event, preview: found.preview } : { status: 'none' });
        }
      })
      .catch(() => {
        if (!cancelled) {
          setState({ status: 'none' });
        }
      });
    return () => {
      cancelled = true;
    };
  }, [client, siteUrl, listRef, isEditMode]);

  return state;
}
