import * as React from 'react';
import styles from './CircleIconList.module.scss';

export interface ICircleIconListProps {
  circleText?: string;
  circleImageUrl?: string;
  title: string;
  /** Up to 5 — extra entries are ignored. */
  items: string[];
}

export default function CircleIconList(props: ICircleIconListProps): React.ReactElement {
  const { circleText, circleImageUrl, title, items } = props;
  const circleStyle = circleImageUrl ? { backgroundImage: `url("${circleImageUrl}")` } : undefined;
  const visible = items.filter(t => !!t && !!t.trim()).slice(0, 5);

  return (
    <div className={styles.card}>
      <div className={styles.circle} style={circleStyle}>
        {!circleImageUrl && circleText}
      </div>
      <p className={styles.title}>{title}</p>
      <div className={styles.list}>
        {visible.map((text, i) => (
          <div className={styles.row} key={i}>
            <p className={styles.rowText}>{text}</p>
          </div>
        ))}
      </div>
    </div>
  );
}
