import * as React from 'react';
import styles from './LogoGridFeature.module.scss';

export interface ILogoItem {
  logoUrl: string;
  linkUrl?: string;
}

export interface ILogoGridFeatureProps {
  imageUrl?: string;
  /** Up to 8 — extra entries are ignored. */
  logos: ILogoItem[];
  title: string;
  bodyText: string;
}

export default function LogoGridFeature(props: ILogoGridFeatureProps): React.ReactElement {
  const { imageUrl, logos, title, bodyText } = props;
  const photoStyle = imageUrl ? { backgroundImage: `url("${imageUrl}")` } : undefined;
  const visibleLogos = logos.filter(l => !!l.logoUrl).slice(0, 8);

  return (
    <div className={styles.wrap}>
      <div className={styles.left}>
        <div className={styles.photo} style={photoStyle} />
        <div className={styles.logos}>
          {visibleLogos.map((logo, i) =>
            logo.linkUrl ? (
              <a className={styles.logo} href={logo.linkUrl} key={i} target="_blank" rel="noopener noreferrer">
                <img src={logo.logoUrl} alt="" />
              </a>
            ) : (
              <div className={styles.logo} key={i}>
                <img src={logo.logoUrl} alt="" />
              </div>
            )
          )}
        </div>
      </div>
      <div className={styles.right}>
        <h3 className={styles.title}>{title}</h3>
        <p className={styles.body}>{bodyText}</p>
      </div>
    </div>
  );
}
