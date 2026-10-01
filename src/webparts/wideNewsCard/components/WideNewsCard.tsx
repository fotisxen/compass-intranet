import * as React from 'react';
import styles from './WideNewsCard.module.scss';
import { NewsCardView } from '../../../shared/components/NewsCard';

export interface IWideNewsCardProps {
  title: string;
  imageUrl?: string;
  dateText: string;
  linkUrl: string;
}

/** Same visual card as the News Card web part, just taller (335 x 781) — with every field entered manually from the property pane instead of pulled from promoted news. */
export default function WideNewsCard(props: IWideNewsCardProps): React.ReactElement {
  const { title, imageUrl, dateText, linkUrl } = props;

  return (
    <div className={styles.wrap}>
      <NewsCardView item={{ title, date: dateText, url: linkUrl, imageUrl }} />
    </div>
  );
}
