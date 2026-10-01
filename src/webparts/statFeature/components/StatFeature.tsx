import * as React from 'react';
import styles from './StatFeature.module.scss';

export interface IStatColumn {
  title: string;
  text: string;
}

export interface IStatFeatureProps {
  bgColor: string;
  eyebrowText?: string;
  eyebrowColor: string;
  titleText: string;
  titleColor: string;
  bodyText: string;
  bodyColor: string;
  imageUrl?: string;
  columnsColor: string;
  /** Dynamic 0-3 — extra entries beyond 3 are ignored. */
  columns: IStatColumn[];
}

export default function StatFeature(props: IStatFeatureProps): React.ReactElement {
  const { bgColor, eyebrowText, eyebrowColor, titleText, titleColor, bodyText, bodyColor, imageUrl, columnsColor, columns } = props;
  const photoStyle = imageUrl ? { backgroundImage: `url("${imageUrl}")` } : undefined;
  const visibleColumns = columns.filter(c => !!c.title || !!c.text).slice(0, 3);

  return (
    <div className={styles.wrap} style={{ backgroundColor: bgColor }}>
      {eyebrowText && <p className={styles.eyebrow} style={{ color: eyebrowColor }}>{eyebrowText}</p>}
      <h2 className={styles.title} style={{ color: titleColor }}>{titleText}</h2>
      <p className={styles.body} style={{ color: bodyColor }}>{bodyText}</p>
      {imageUrl && <div className={styles.photo} style={photoStyle} />}
      {visibleColumns.length > 0 && (
        <>
          <hr className={styles.divider} />
          <div className={styles.columns}>
            {visibleColumns.map((c, i) => (
              <div className={styles.column} key={i}>
                <p className={styles.columnTitle} style={{ color: columnsColor }}>{c.title}</p>
                <p className={styles.columnText} style={{ color: columnsColor }}>{c.text}</p>
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );
}
