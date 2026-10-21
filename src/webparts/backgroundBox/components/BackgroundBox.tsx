import * as React from 'react';
import styles from './BackgroundBox.module.scss';

export interface IBackgroundBoxProps {
  /** column = inside the column's padding; columnBleed = edge to edge of the column; full = the whole page width; custom = widthValue in widthUnit. */
  widthMode: 'column' | 'columnBleed' | 'full' | 'custom';
  widthValue: number;
  widthUnit: 'px' | '%';
  align: 'left' | 'center' | 'right';
  heightValue: number;
  heightUnit: 'px' | 'vh';
  color: string;
  useGradient: boolean;
  colorEnd: string;
  gradientAngle: number;
  /** 0–100. */
  opacity: number;
  radius: number;
  perCorner: boolean;
  radiusTopLeft: number;
  radiusTopRight: number;
  radiusBottomRight: number;
  radiusBottomLeft: number;
  borderWidth: number;
  borderColor: string;
  shadow: boolean;
  /** Pixels the content below is pulled up over the box (negative bottom margin). */
  pullUp: number;
  isEditMode: boolean;
}

/** SharePoint's padding around every web part, which "edge to edge" cancels. */
const ZONE_PADDING = 16;

/** Where a full-page-width box has to start and how wide it is, measured from the box's own column. */
function useFullWidth(active: boolean, ref: React.RefObject<HTMLDivElement>): { width: number; left: number } | undefined {
  const [full, setFull] = React.useState<{ width: number; left: number } | undefined>();

  React.useLayoutEffect(() => {
    if (!active) {
      setFull(undefined);
      return undefined;
    }
    const update = (): void => {
      const host = ref.current && ref.current.parentElement;
      if (!host) {
        return;
      }
      // clientWidth (not 100vw) so a scrollbar is not counted as page width.
      setFull({ width: document.documentElement.clientWidth, left: -host.getBoundingClientRect().left });
    };
    update();
    window.addEventListener('resize', update);
    const observer = typeof ResizeObserver === 'function' ? new ResizeObserver(update) : undefined;
    if (observer && ref.current && ref.current.parentElement) {
      observer.observe(ref.current.parentElement);
    }
    return () => {
      window.removeEventListener('resize', update);
      if (observer) {
        observer.disconnect();
      }
    };
  }, [active, ref]);

  return full;
}

export default function BackgroundBox(props: IBackgroundBoxProps): React.ReactElement {
  const ref = React.useRef<HTMLDivElement>(null);
  const full = useFullWidth(props.widthMode === 'full', ref);

  const style: React.CSSProperties = {
    height: `${props.heightValue}${props.heightUnit}`,
    background: props.useGradient
      ? `linear-gradient(${props.gradientAngle}deg, ${props.color}, ${props.colorEnd})`
      : props.color,
    // one CSS opacity on the whole box (the colour keeps its picked value)
    opacity: Math.max(0, Math.min(100, props.opacity)) / 100,
    borderRadius: props.perCorner
      ? `${props.radiusTopLeft}px ${props.radiusTopRight}px ${props.radiusBottomRight}px ${props.radiusBottomLeft}px`
      : `${props.radius}px`,
    border: props.borderWidth > 0 ? `${props.borderWidth}px solid ${props.borderColor}` : undefined,
    boxShadow: props.shadow ? '0 8px 28px rgba(8, 34, 68, 0.25)' : undefined,
    marginBottom: props.pullUp > 0 ? -props.pullUp : undefined
  };

  if (props.widthMode === 'columnBleed') {
    style.width = `calc(100% + ${ZONE_PADDING * 2}px)`;
    style.marginLeft = -ZONE_PADDING;
  } else if (props.widthMode === 'full') {
    // Until it has been measured, fill the column so there is no jump.
    style.width = full ? full.width : '100%';
    style.marginLeft = full ? full.left : 0;
  } else if (props.widthMode === 'custom') {
    style.width = `${props.widthValue}${props.widthUnit}`;
    style.maxWidth = '100%';
    style.marginLeft = props.align === 'left' ? 0 : 'auto';
    style.marginRight = props.align === 'right' ? 0 : props.align === 'center' ? 'auto' : undefined;
  } else {
    style.width = '100%';
  }

  return (
    <div ref={ref} className={styles.box} style={style}>
      {props.isEditMode && <span className={styles.hint}>Background box — {props.widthMode === 'full' ? 'full width' : props.widthMode === 'custom' ? `${props.widthValue}${props.widthUnit}` : 'column'} × {props.heightValue}{props.heightUnit}</span>}
    </div>
  );
}
