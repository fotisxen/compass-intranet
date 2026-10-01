import * as React from 'react';
import styles from './HighlightBanner.module.scss';

export interface IHighlightBannerProps {
  title: string;
  titleColor: string;
  items: string[];
  bodyColor: string;
  bgColor: string;
}

export default function HighlightBanner(props: IHighlightBannerProps): React.ReactElement {
  const { title, titleColor, items, bodyColor, bgColor } = props;
  const visible = items.filter(t => !!t && !!t.trim());

  return (
    <div className={styles.banner} style={{ backgroundColor: bgColor }}>
      <p className={styles.title} style={{ color: titleColor }}>{title}</p>
      <ul className={styles.body} style={{ color: bodyColor }}>
        {visible.map((text, i) => <li key={i}>{text}</li>)}
      </ul>
    </div>
  );
}
