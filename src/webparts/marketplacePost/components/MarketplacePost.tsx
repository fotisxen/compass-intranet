import * as React from 'react';
import * as ReactDom from 'react-dom';
import type { SPHttpClient } from '@microsoft/sp-http';
import shared from '../../../shared/marketplace/Marketplace.module.scss';
import styles from './MarketplacePost.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';
import { resizePhoto, type IResizedPhoto } from '../../../shared/marketplace/imageResize';
import { announcePosted } from '../../../shared/marketplace/marketplaceBus';
import { ensureMarketplaceLists, loadCategories, submitListing, type ICategory } from '../../../shared/marketplace/marketplaceApi';

export interface IMarketplacePostProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  canManageLists: boolean;
  eyebrow: string;
  heading: string;
  buttonLabel: string;
  maxPhotos: number;
}

interface IFormProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  canManageLists: boolean;
  maxPhotos: number;
  onClose: () => void;
}

function PostForm(props: IFormProps): React.ReactElement {
  const { spHttpClient, siteUrl, canManageLists, maxPhotos, onClose } = props;
  const [categories, setCategories] = React.useState<ICategory[]>([]);
  const [listsMissing, setListsMissing] = React.useState(false);
  const [title, setTitle] = React.useState('');
  const [category, setCategory] = React.useState('');
  const [price, setPrice] = React.useState('');
  const [area, setArea] = React.useState('');
  const [details, setDetails] = React.useState('');
  const [photos, setPhotos] = React.useState<IResizedPhoto[]>([]);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>();
  const [done, setDone] = React.useState<{ failedPhotos: number } | undefined>();
  const fileInput = React.useRef<HTMLInputElement>(null);

  React.useEffect(() => {
    loadCategories(spHttpClient, siteUrl)
      .then(result => {
        if (result === 'missing') {
          setListsMissing(true);
        } else if (typeof result !== 'string') {
          setCategories(result.data);
        }
      })
      .catch(() => setError('The categories could not be loaded.'));
  }, [spHttpClient, siteUrl]);

  React.useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape' && !busy) {
        onClose();
      }
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onClose, busy]);

  // Release the preview thumbnails when the form closes.
  const photosRef = React.useRef(photos);
  photosRef.current = photos;
  React.useEffect(() => () => photosRef.current.forEach(p => URL.revokeObjectURL(p.previewUrl)), []);

  const addFiles = async (files: FileList | null): Promise<void> => {
    if (!files) {
      return;
    }
    setError(undefined);
    const room = maxPhotos - photos.length;
    const picked = Array.prototype.slice.call(files, 0, Math.max(0, room)) as File[];
    const added: IResizedPhoto[] = [];
    for (let i = 0; i < picked.length; i++) {
      try {
        added.push(await resizePhoto(picked[i], photos.length + i));
      } catch (e) {
        setError((e as Error).message);
      }
    }
    setPhotos(prev => [...prev, ...added]);
    if (files.length > room) {
      setError(`Only ${maxPhotos} photos fit in one listing — the extra ones were left out.`);
    }
    if (fileInput.current) {
      fileInput.current.value = '';
    }
  };

  const removePhoto = (index: number): void => {
    setPhotos(prev => {
      URL.revokeObjectURL(prev[index].previewUrl);
      return prev.filter((_, i) => i !== index);
    });
  };

  const setup = async (): Promise<void> => {
    setBusy(true);
    try {
      await ensureMarketplaceLists(spHttpClient, siteUrl);
      const result = await loadCategories(spHttpClient, siteUrl);
      if (typeof result !== 'string') {
        setCategories(result.data);
        setListsMissing(false);
      }
    } catch (e) {
      setError((e as Error).message);
    }
    setBusy(false);
  };

  const submit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    if (!title.trim() || !category) {
      setError('Please add a title and choose a category.');
      return;
    }
    setBusy(true);
    setError(undefined);
    try {
      const result = await submitListing(spHttpClient, siteUrl, {
        title: title.trim(),
        category,
        price: price.trim(),
        area: area.trim(),
        details: details.trim(),
        photos: photos.map(p => ({ fileName: p.fileName, data: p.data }))
      });
      announcePosted();
      setDone(result);
    } catch (err) {
      const message = (err as Error).message || '';
      setError(/\(40[13]\)/.test(message)
        ? 'You do not have permission to post yet. Please contact HR or IT.'
        : 'Your listing could not be sent. Please try again.');
    }
    setBusy(false);
  };

  return (
    <div className={shared.overlay} onClick={busy ? undefined : onClose} role="presentation">
      <div className={`${shared.modal} ${styles.formModal}`} role="dialog" aria-modal="true" aria-label="Post an item" onClick={e => e.stopPropagation()}>
        <button type="button" className={shared.close} onClick={onClose} disabled={busy} aria-label="Close">✕</button>
        {done ? (
          <div className={styles.done}>
            <h3 className={styles.doneTitle}>Thank you!</h3>
            <p className={styles.doneText}>
              Your listing was sent to HR for approval. It will appear on the Marketplace as soon as it is approved.
              {done.failedPhotos > 0 ? ` ${done.failedPhotos} photo(s) could not be uploaded — you can add them later from "My listings".` : ''}
            </p>
            <button type="button" className={styles.submit} onClick={onClose}>Close</button>
          </div>
        ) : listsMissing ? (
          <div className={styles.done}>
            <p className={styles.doneText}>
              {canManageLists ? 'The Marketplace lists have not been created on this site yet.' : 'The Marketplace is not set up yet. Please ask a site owner to open this page once.'}
            </p>
            {canManageLists && <button type="button" className={styles.submit} onClick={() => { setup().catch(() => undefined); }} disabled={busy}>Create the Marketplace lists</button>}
            {error && <p className={styles.error}>{error}</p>}
          </div>
        ) : (
          <form onSubmit={e => { submit(e).catch(() => undefined); }}>
            <h3 className={styles.formTitle}>Post an item</h3>

            <div className={styles.field}>
              <span className={styles.label}>PHOTOS</span>
              <div className={styles.photos}>
                {photos.map((p, i) => (
                  <div className={styles.thumb} key={p.previewUrl}>
                    <img src={p.previewUrl} alt="" />
                    <button type="button" className={styles.thumbRemove} onClick={() => removePhoto(i)} aria-label="Remove photo">✕</button>
                  </div>
                ))}
                {photos.length < maxPhotos && (
                  <button type="button" className={styles.photoAdd} onClick={() => fileInput.current && fileInput.current.click()} disabled={busy}>
                    <span style={{ fontSize: '1.6em', lineHeight: 1 }}>+</span>
                    <span>Add photo</span>
                  </button>
                )}
              </div>
              <input ref={fileInput} type="file" accept="image/*" multiple={true} hidden={true} onChange={e => { addFiles(e.target.files).catch(() => undefined); }} />
              <span className={styles.hint}>Up to {maxPhotos} photos ({photos.length}/{maxPhotos}). They are shrunk automatically.</span>
            </div>

            <label className={styles.field}>
              <span className={styles.label}>TITLE</span>
              <input className={styles.input} type="text" value={title} onChange={e => setTitle(e.target.value)} maxLength={120} required={true} />
            </label>

            <label className={styles.field}>
              <span className={styles.label}>CATEGORY</span>
              <select className={styles.select} value={category} onChange={e => setCategory(e.target.value)} required={true}>
                <option value="" disabled={true}>Choose a category</option>
                {categories.map(c => <option key={c.id} value={c.name}>{c.name}</option>)}
              </select>
            </label>

            <label className={styles.field}>
              <span className={styles.label}>COST</span>
              <input className={styles.input} type="text" value={price} onChange={e => setPrice(e.target.value)} maxLength={100} placeholder="e.g. €20, Free, Make an offer" />
            </label>

            <label className={styles.field}>
              <span className={styles.label}>AREA</span>
              <input className={styles.input} type="text" value={area} onChange={e => setArea(e.target.value)} maxLength={100} placeholder="e.g. Athens HQ" />
            </label>

            <label className={styles.field}>
              <span className={styles.label}>DESCRIPTION</span>
              <textarea className={styles.textarea} value={details} onChange={e => setDetails(e.target.value)} maxLength={2000} />
            </label>

            {error && <p className={styles.error}>{error}</p>}
            <div className={styles.actions}>
              <button type="button" className={styles.cancel} onClick={onClose} disabled={busy}>Cancel</button>
              <button type="submit" className={styles.submit} disabled={busy}>{busy ? 'Sending…' : 'Send for approval'}</button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
}

export default function MarketplacePost(props: IMarketplacePostProps): React.ReactElement {
  const [open, setOpen] = React.useState(false);

  React.useEffect(() => {
    ensureInterFont();
  }, []);

  return (
    <div className={shared.page}>
      <div className={styles.card}>
        {props.eyebrow && <span className={styles.eyebrow}>{props.eyebrow}</span>}
        {props.heading && <p className={styles.heading}>{props.heading}</p>}
        <button type="button" className={styles.postButton} onClick={() => setOpen(true)}>
          <span aria-hidden="true">+</span> {props.buttonLabel}
        </button>
      </div>
      {open && ReactDom.createPortal(
        <PostForm
          spHttpClient={props.spHttpClient}
          siteUrl={props.siteUrl}
          canManageLists={props.canManageLists}
          maxPhotos={props.maxPhotos}
          onClose={() => setOpen(false)}
        />,
        document.body
      )}
    </div>
  );
}
