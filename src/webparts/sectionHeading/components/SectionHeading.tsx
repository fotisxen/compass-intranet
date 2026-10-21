import * as React from 'react';
import styles from './SectionHeading.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';

export interface ISectionHeadingProps {
  title: string;
  linkText: string;
  linkUrl: string;
  showArrow: boolean;
  newTab: boolean;
  /** Pixels at the 1440px reference width (scales with the page). */
  titleSize: number;
  isEditMode?: boolean;
}

export default function SectionHeading(props: ISectionHeadingProps): React.ReactElement {
  React.useEffect(() => {
    ensureInterFont();
  }, []);

  const hasLink = !!props.linkText && !!props.linkUrl;
  return (
    <div className={styles.page}>
      <div className={styles.row}>
        <h2 className={styles.title} style={{ fontSize: `${props.titleSize / 20}em` }}>{props.title}</h2>
        {!hasLink && !!props.linkText && props.isEditMode && (
          <span className={styles.link} title="Set the link address in this web part's settings">{props.linkText}{props.showArrow ? ' →' : ''} (no link address yet)</span>
        )}
        {hasLink && (
          <a
            className={styles.link}
            href={props.linkUrl}
            target={props.newTab ? '_blank' : undefined}
            rel={props.newTab ? 'noopener noreferrer' : undefined}
          >
            {props.linkText}{props.showArrow ? ' →' : ''}
          </a>
        )}
      </div>
    </div>
  );
}
