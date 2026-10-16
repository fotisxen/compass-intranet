/**
 * The Marketplace web parts are separate bundles on one page, so they can't
 * share JavaScript state directly. They share it through `window` instead: the
 * filter web part writes the chosen category and search text, the listings web
 * part listens; the post web part announces a new post so "My listings" reloads.
 */
export const ALL_CATEGORIES = 'All';

export interface IMarketplaceFilter {
  category: string;
  search: string;
}

const FILTER_EVENT = 'compass-marketplace-filter';
const POSTED_EVENT = 'compass-marketplace-posted';
const GLOBAL_KEY = '__compassMarketplaceFilter';

interface IWindowWithFilter {
  [GLOBAL_KEY]?: IMarketplaceFilter;
}

export function getFilter(): IMarketplaceFilter {
  return (window as unknown as IWindowWithFilter)[GLOBAL_KEY] || { category: ALL_CATEGORIES, search: '' };
}

export function setFilter(partial: Partial<IMarketplaceFilter>): void {
  const next = { ...getFilter(), ...partial };
  (window as unknown as IWindowWithFilter)[GLOBAL_KEY] = next;
  window.dispatchEvent(new CustomEvent<IMarketplaceFilter>(FILTER_EVENT, { detail: next }));
}

/** Calls back whenever the category or the search text changes. Returns the function that stops listening. */
export function onFilterChange(callback: (filter: IMarketplaceFilter) => void): () => void {
  const handler = (e: Event): void => callback((e as CustomEvent<IMarketplaceFilter>).detail);
  window.addEventListener(FILTER_EVENT, handler);
  return () => window.removeEventListener(FILTER_EVENT, handler);
}

export function announcePosted(): void {
  window.dispatchEvent(new CustomEvent(POSTED_EVENT));
}

export function onPosted(callback: () => void): () => void {
  window.addEventListener(POSTED_EVENT, callback);
  return () => window.removeEventListener(POSTED_EVENT, callback);
}
