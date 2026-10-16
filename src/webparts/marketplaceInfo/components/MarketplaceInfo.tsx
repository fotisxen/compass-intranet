import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import shared from '../../../shared/marketplace/Marketplace.module.scss';
import styles from './MarketplaceInfo.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import { getList } from '../../../shared/components/spListSetup';
import { onPosted } from '../../../shared/marketplace/marketplaceBus';
import {
  LISTINGS_LIST_TITLE,
  STATUS_APPROVED,
  STATUS_REJECTED,
  loadMyListings,
  type IListing
} from '../../../shared/marketplace/marketplaceApi';

export interface IMarketplaceInfoProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  isEditMode: boolean;
  rulesTitle: string;
  rulesText: string;
  contactLabel: string;
  contactEmail: string;
  contactSubject: string;
  contactBody: string;
  manageUrl: string;
  expiryDays: number;
}

const MY_LISTINGS_SHOWN = 5;

function statusOf(listing: IListing, expiryDays: number): { text: string; className: string } {
  if (listing.status === STATUS_REJECTED) {
    return { text: 'Rejected', className: styles.statusRejected };
  }
  if (listing.status !== STATUS_APPROVED) {
    return { text: 'Pending approval', className: styles.statusPending };
  }
  const expired = expiryDays > 0 && Date.now() - listing.created.getTime() > expiryDays * 24 * 60 * 60 * 1000;
  return expired ? { text: 'Expired', className: '' } : { text: 'Live', className: styles.statusLive };
}

function MyListings(props: IMarketplaceInfoProps): React.ReactElement {
  const { spHttpClient, siteUrl, expiryDays, manageUrl } = props;
  const [items, setItems] = React.useState<IListing[]>([]);
  const [manageLink, setManageLink] = React.useState<string>(manageUrl);

  const load = React.useCallback(async (): Promise<void> => {
    const result = await loadMyListings(spHttpClient, siteUrl, MY_LISTINGS_SHOWN);
    setItems(typeof result === 'string' ? [] : result.data);
  }, [spHttpClient, siteUrl]);

  React.useEffect(() => {
    load().catch(() => setItems([]));
    return onPosted(() => { load().catch(() => undefined); });
  }, [load]);

  React.useEffect(() => {
    if (manageUrl) {
      setManageLink(manageUrl);
      return;
    }
    getList(spHttpClient, siteUrl, LISTINGS_LIST_TITLE)
      .then(info => setManageLink(info && info.DefaultViewUrl ? info.DefaultViewUrl : ''))
      .catch(() => setManageLink(''));
  }, [spHttpClient, siteUrl, manageUrl]);

  return (
    <div className={styles.col}>
      <h4 className={styles.heading}>MY LISTINGS</h4>
      {items.length === 0 ? (
        <div className={styles.empty}>You have not posted anything yet.</div>
      ) : (
        <ul className={styles.myList}>
          {items.map(item => {
            const status = statusOf(item, expiryDays);
            return (
              <li className={styles.myItem} key={item.id}>
                <span className={styles.myTitle} title={item.title}>{item.title}</span>
                <span className={`${styles.myStatus} ${status.className}`}>{status.text}</span>
              </li>
            );
          })}
        </ul>
      )}
      {manageLink && <a className={styles.manage} href={manageLink}>Manage listings →</a>}
    </div>
  );
}

export default function MarketplaceInfo(props: IMarketplaceInfoProps): React.ReactElement {
  const { rulesTitle, rulesText, contactLabel, contactEmail, contactSubject, contactBody, isEditMode } = props;

  React.useEffect(() => {
    ensureInterFont();
  }, []);

  const rules = rulesText.split(/\r?\n/).map(r => r.trim()).filter(r => !!r);
  const mailto = contactEmail
    ? `mailto:${contactEmail}?subject=${encodeURIComponent(contactSubject)}&body=${encodeURIComponent(contactBody)}`
    : '';

  return (
    <div className={shared.page}>
      <div className={styles.row}>
        <MyListings {...props} />

        <div className={styles.col}>
          {rulesTitle && <h4 className={styles.heading}>{rulesTitle}</h4>}
          <ul className={styles.rules}>
            {rules.map(r => <li key={r}>{r}</li>)}
          </ul>
        </div>

        <div className={styles.contactCol}>
          {mailto ? (
            <a className={styles.contact} href={mailto}>{contactLabel} →</a>
          ) : isEditMode ? (
            <span className={styles.editHint}>Set the HR email in the settings of this web part to show the contact link.</span>
          ) : null}
        </div>
      </div>
    </div>
  );
}
