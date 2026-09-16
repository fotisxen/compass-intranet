import * as React from 'react';
import styles from './PersonSpotlightCard.module.scss';
import type { IPersonSpotlight } from './data/mockData';

export interface IPersonSpotlightCardProps {
  heading: string;
  background: string;
  /** Hand-authored list (see data/mockData.ts) — the arrows cycle through these one at a time. */
  people: IPersonSpotlight[];
}

export interface IPersonSpotlightCardState {
  index: number;
}

/**
 * Long employee names/titles must stay on one line, side by side with the
 * details column (never wrap below it), so instead of truncating or
 * wrapping, the font shrinks based on how long the text actually is.
 */
function fitFontSize(text: string, base: number, min: number, startShrinkingAt: number, pxPerChar: number): number {
  if (text.length <= startShrinkingAt) {
    return base;
  }
  const size = base - (text.length - startShrinkingAt) * pxPerChar;
  return Math.max(min, Math.round(size * 10) / 10);
}

export default class PersonSpotlightCard extends React.Component<IPersonSpotlightCardProps, IPersonSpotlightCardState> {
  constructor(props: IPersonSpotlightCardProps) {
    super(props);
    this.state = { index: 0 };
  }

  private _prev = (): void => {
    this.setState(prev => ({ index: (prev.index - 1 + this.props.people.length) % this.props.people.length }));
  };

  private _next = (): void => {
    this.setState(prev => ({ index: (prev.index + 1) % this.props.people.length }));
  };

  public render(): React.ReactElement<IPersonSpotlightCardProps> {
    const { heading, background, people } = this.props;
    const person = people[this.state.index];

    const nameFontSize = fitFontSize(person.name, 14, 11, 16, 0.3);
    const titleFontSize = fitFontSize(person.title, 11, 9, 20, 0.15);
    const newTitleFontSize = person.newTitle ? fitFontSize(person.newTitle, 11, 9, 18, 0.15) : undefined;

    return (
      <div className={styles.card} style={{ background }}>
        {people.length > 1 && (
          <button className={styles.arrowButton} onClick={this._prev} aria-label="Previous">←</button>
        )}

        <p className={styles.heading}>{heading}</p>

        <span className={styles.avatar}>{person.initials}</span>
        <div className={styles.nameBlock}>
          <p className={styles.name} style={{ fontSize: nameFontSize }}>{person.name}</p>
          <p className={styles.title} style={{ fontSize: titleFontSize }}>{person.title}</p>
          {person.newTitle && (
            <p className={styles.newTitle} style={{ fontSize: newTitleFontSize }}>→ {person.newTitle}</p>
          )}
        </div>

        <div className={styles.details}>
          {person.fields.map(f => (
            <div className={styles.field} key={f.label}>
              <span className={styles.fieldLabel}>{f.label}</span>
              <span className={styles.fieldValue} style={{ fontSize: fitFontSize(f.value, 11, 9, 12, 0.2) }}>
                {f.value}
              </span>
            </div>
          ))}
        </div>

        {people.length > 1 && (
          <button className={styles.arrowButton} onClick={this._next} aria-label="Next">→</button>
        )}
      </div>
    );
  }
}
