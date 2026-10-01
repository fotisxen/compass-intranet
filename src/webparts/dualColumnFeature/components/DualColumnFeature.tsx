import * as React from 'react';
import styles from './DualColumnFeature.module.scss';

export interface IDualColumnFeatureProps {
  bgColor: string;
  leftTitle: string;
  leftItems: string[];
  rightTitle: string;
  rightItems: string[];
}

export default function DualColumnFeature(props: IDualColumnFeatureProps): React.ReactElement {
  const { bgColor, leftTitle, leftItems, rightTitle, rightItems } = props;
  const left = leftItems.filter(t => !!t && !!t.trim());
  const right = rightItems.filter(t => !!t && !!t.trim());

  return (
    <div className={styles.wrap} style={{ backgroundColor: bgColor }}>
      <div className={styles.card}>
        <p className={styles.title}>{leftTitle}</p>
        <ul className={styles.body}>
          {left.map((text, i) => <li key={i}>{text}</li>)}
        </ul>
      </div>
      <div className={styles.card}>
        <p className={styles.title}>{rightTitle}</p>
        <ul className={styles.body}>
          {right.map((text, i) => <li key={i}>{text}</li>)}
        </ul>
      </div>
    </div>
  );
}
