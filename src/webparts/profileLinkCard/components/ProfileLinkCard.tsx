import * as React from 'react';
import styles from './ProfileLinkCard.module.scss';

export interface IProfileLinkCardProps {
  topText: string;
  topLinkUrl?: string;
  bottomText?: string;
}

export default function ProfileLinkCard(props: IProfileLinkCardProps): React.ReactElement {
  const { topText, topLinkUrl, bottomText } = props;

  return (
    <div className={styles.card}>
      {topLinkUrl ? (
        <a className={styles.topLink} href={topLinkUrl}>{topText}</a>
      ) : (
        <p className={styles.topText}>{topText}</p>
      )}
      {bottomText && <p className={styles.bottomText}>{bottomText}</p>}
    </div>
  );
}
