import * as React from 'react';
import styles from './FeatureSplit.module.scss';

export interface IFeatureSplitProps {
  title: string;
  titleColor: string;
  bodyText: string;
  bodyColor: string;
  imageUrl?: string;
}

export default function FeatureSplit(props: IFeatureSplitProps): React.ReactElement {
  const { title, titleColor, bodyText, bodyColor, imageUrl } = props;
  const photoStyle = imageUrl ? { backgroundImage: `url("${imageUrl}")` } : undefined;

  return (
    <div className={styles.wrap}>
      <div className={styles.text}>
        <hr className={styles.divider} />
        <h3 className={styles.title} style={{ color: titleColor }}>{title}</h3>
        <p className={styles.body} style={{ color: bodyColor }}>{bodyText}</p>
      </div>
      <div className={styles.photo} style={photoStyle} />
    </div>
  );
}
