import * as React from 'react';
import styles from './TestimonialCarousel.module.scss';
import ArrowIcon from '../../../shared/components/ArrowIcon';

export interface ITestimonial {
  quote: string;
  author: string;
}

export interface ITestimonialCarouselProps {
  title: string;
  titleColor: string;
  bgColor: string;
  testimonials: ITestimonial[];
}

// One card's width (see .cardItem) plus the track's own gap.
const STEP = 254 + 16;

export default function TestimonialCarousel(props: ITestimonialCarouselProps): React.ReactElement {
  const { title, titleColor, bgColor, testimonials } = props;
  const trackRef = React.useRef<HTMLDivElement>(null);
  const [atStart, setAtStart] = React.useState(true);
  const [atEnd, setAtEnd] = React.useState(false);

  const updateEdges = React.useCallback(() => {
    const el = trackRef.current;
    if (!el) {
      return;
    }
    setAtStart(el.scrollLeft <= 0);
    setAtEnd(el.scrollLeft + el.clientWidth >= el.scrollWidth - 1);
  }, []);

  React.useEffect(() => {
    updateEdges();
  }, [updateEdges, testimonials.length]);

  const scrollByStep = (direction: 1 | -1): void => {
    trackRef.current?.scrollBy({ left: direction * STEP, behavior: 'smooth' });
  };

  // A normal (vertical) mouse wheel does nothing on a horizontal-only track
  // by default — redirect it to horizontal scrolling so hovering the
  // carousel and just scrolling the wheel works, no drag or shift needed.
  const onWheel = (e: React.WheelEvent<HTMLDivElement>): void => {
    if (Math.abs(e.deltaY) <= Math.abs(e.deltaX)) {
      return;
    }
    e.currentTarget.scrollLeft += e.deltaY;
    e.preventDefault();
  };

  const canScroll = testimonials.length > 0;

  return (
    <div className={styles.wrap} style={{ backgroundColor: bgColor }}>
      <h3 className={styles.title} style={{ color: titleColor }}>{title}</h3>
      <div className={styles.track} ref={trackRef} onScroll={updateEdges} onWheel={onWheel}>
        {testimonials.map((t, i) => (
          <div className={styles.cardItem} key={`${t.author}-${i}`}>
            <p className={styles.quoteText}>{t.quote}</p>
            <p className={styles.author}>{t.author}</p>
          </div>
        ))}
      </div>
      {canScroll && (
        <div className={styles.arrowsRow}>
          <button
            className={styles.arrow}
            onClick={() => scrollByStep(-1)}
            disabled={atStart}
            aria-label="Previous"
          >
            <ArrowIcon direction="left" />
          </button>
          <button
            className={styles.arrow}
            onClick={() => scrollByStep(1)}
            disabled={atEnd}
            aria-label="Next"
          >
            <ArrowIcon direction="right" />
          </button>
        </div>
      )}
    </div>
  );
}
