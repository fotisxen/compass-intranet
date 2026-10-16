import * as React from 'react';
import * as ReactDom from 'react-dom';
import styles from './Marketplace.module.scss';
import ArrowIcon from '../components/ArrowIcon';
import { type IListing, initialsOf, shortName, teamsChatUrl, timeAgo } from './marketplaceApi';

// Soft badge colours; "Free" is always green, the other categories get a stable colour from their name.
const BADGE_COLOURS: { bg: string; fg: string }[] = [
  { bg: '#e2ebfb', fg: '#1d4d9a' },
  { bg: '#fdf0d5', fg: '#8a5a00' },
  { bg: '#ebe5fb', fg: '#52399b' },
  { bg: '#fbe3ea', fg: '#9b2c52' },
  { bg: '#d9f1ee', fg: '#0f6b63' }
];

function badgeColours(category: string): { bg: string; fg: string } {
  if (category.trim().toLowerCase() === 'free') {
    return { bg: '#d9f5e2', fg: '#17653a' };
  }
  let hash = 0;
  for (let i = 0; i < category.length; i++) {
    hash = (hash * 31 + category.charCodeAt(i)) >>> 0;
  }
  return BADGE_COLOURS[hash % BADGE_COLOURS.length];
}

function PlaceholderIcon(): React.ReactElement {
  return (
    <svg className={styles.noPhoto} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <rect x="3" y="4" width="18" height="16" rx="2" />
      <circle cx="9" cy="10" r="1.6" />
      <path d="M21 16l-5-5-8 8" />
    </svg>
  );
}

interface IPhotoCarouselProps {
  photos: string[];
  alt: string;
  category: string;
  large?: boolean;
}

/** The photo area: category badge top-left, arrows (and dots) when there is more than one photo. */
export function PhotoCarousel(props: IPhotoCarouselProps): React.ReactElement {
  const { photos, alt, category, large } = props;
  const [index, setIndex] = React.useState(0);
  const count = photos.length;
  const colours = badgeColours(category);

  const go = (e: React.MouseEvent, step: number): void => {
    // Arrows must not also open the listing the card belongs to.
    e.stopPropagation();
    setIndex(i => (i + step + count) % count);
  };

  return (
    <div className={`${styles.photo} ${large ? styles.photoLarge : ''}`}>
      {count > 0 ? <img src={photos[index % count]} alt={alt} loading="lazy" /> : <PlaceholderIcon />}
      {category && (
        <span className={styles.badge} style={{ background: colours.bg, color: colours.fg }}>{category}</span>
      )}
      {count > 1 && (
        <>
          <button type="button" className={`${styles.arrow} ${styles.arrowLeft}`} onClick={e => go(e, -1)} aria-label="Previous photo"><ArrowIcon direction="left" /></button>
          <button type="button" className={`${styles.arrow} ${styles.arrowRight}`} onClick={e => go(e, 1)} aria-label="Next photo"><ArrowIcon direction="right" /></button>
          <div className={styles.dots} aria-hidden="true">
            {photos.map((p, i) => <span key={p} className={`${styles.dot} ${i === index % count ? styles.dotActive : ''}`} />)}
          </div>
        </>
      )}
    </div>
  );
}

function SellerAvatar(props: { name: string; email: string }): React.ReactElement {
  const { name, email } = props;
  const [failed, setFailed] = React.useState(false);
  const photo = email && !failed ? `/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(email)}` : undefined;
  return (
    <span className={styles.avatar} aria-hidden="true">
      {initialsOf(name)}
      {photo && <img src={photo} alt="" onError={() => setFailed(true)} />}
    </span>
  );
}

function Seller(props: { listing: IListing }): React.ReactElement {
  const { listing } = props;
  return (
    <span className={styles.seller}>
      <SellerAvatar name={listing.authorName} email={listing.authorEmail} />
      <span className={styles.sellerName}>{shortName(listing.authorName)}</span>
    </span>
  );
}

function ChatButton(props: { listing: IListing }): React.ReactElement {
  const { listing } = props;
  if (!listing.authorEmail) {
    return <></>;
  }
  return (
    <a
      className={styles.pill}
      href={teamsChatUrl(listing.authorEmail)}
      target="_blank"
      rel="noopener noreferrer"
      onClick={e => e.stopPropagation()}
    >
      Chat on Teams
    </a>
  );
}

function metaLine(listing: IListing): string {
  return [listing.area, timeAgo(listing.created)].filter(p => !!p).join(' · ');
}

/** One listing card; clicking anywhere on it (except the arrows and Chat button) opens it. */
export function ListingCard(props: { listing: IListing; onOpen: (listing: IListing) => void }): React.ReactElement {
  const { listing, onOpen } = props;
  const open = (): void => onOpen(listing);
  return (
    <div
      className={styles.card}
      role="button"
      tabIndex={0}
      onClick={open}
      onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); open(); } }}
    >
      <PhotoCarousel photos={props.listing.photos} alt={listing.title} category={listing.category} />
      <div className={styles.body}>
        <h4 className={styles.title}>{listing.title}</h4>
        {listing.price && <p className={styles.price}>{listing.price}</p>}
        <p className={styles.meta}>{metaLine(listing)}</p>
      </div>
      <div className={styles.footer}>
        <Seller listing={listing} />
        <ChatButton listing={listing} />
      </div>
    </div>
  );
}

/** The big popup with every photo, the full text and the seller; blurs what is behind it. */
export function ListingModal(props: { listing: IListing; onClose: () => void }): React.ReactElement {
  const { listing, onClose } = props;

  React.useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') {
        onClose();
      }
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onClose]);

  return ReactDom.createPortal(
    <div className={styles.overlay} onClick={onClose} role="presentation">
      <div className={styles.modal} role="dialog" aria-modal="true" aria-label={listing.title} onClick={e => e.stopPropagation()}>
        <button type="button" className={styles.close} onClick={onClose} aria-label="Close">✕</button>
        <PhotoCarousel photos={listing.photos} alt={listing.title} category={listing.category} large={true} />
        <div className={styles.modalBody}>
          <h3 className={styles.modalTitle}>{listing.title}</h3>
          {listing.price && <p className={styles.modalPrice}>{listing.price}</p>}
          <p className={styles.modalMeta}>{metaLine(listing)}</p>
          {listing.details && <p className={styles.modalText}>{listing.details}</p>}
        </div>
        <div className={styles.footer}>
          <Seller listing={listing} />
          <ChatButton listing={listing} />
        </div>
      </div>
    </div>,
    document.body
  );
}
