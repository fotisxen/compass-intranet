import * as React from 'react';
import styles from './LinkButton.module.scss';
import { ensureInterFont } from '../../../shared/components/ensureFonts';

export interface ILinkButtonProps {
  label: string;
  url: string;
  showArrow: boolean;
  newTab: boolean;
  color: string;
  textColor: string;
  align: 'left' | 'center' | 'right';
  /** Space above the button, in pixels. */
  spaceAbove: number;
}

export default function LinkButton(props: ILinkButtonProps): React.ReactElement {
  React.useEffect(() => {
    ensureInterFont();
  }, []);

  return (
    <div style={{ paddingTop: props.spaceAbove }} className={`${styles.page} ${props.align === 'center' ? styles.center : props.align === 'right' ? styles.right : styles.left}`}>
      <a
        className={styles.button}
        href={props.url || undefined}
        target={props.newTab ? '_blank' : undefined}
        rel={props.newTab ? 'noopener noreferrer' : undefined}
        style={{ background: props.color, color: props.textColor }}
      >
        {props.label}{props.showArrow ? ' →' : ''}
      </a>
    </div>
  );
}
