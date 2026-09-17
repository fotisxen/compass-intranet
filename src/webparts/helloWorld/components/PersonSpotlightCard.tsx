import * as React from 'react';
import styles from './PersonSpotlightCard.module.scss';
import type { IPersonSpotlight } from './data/mockData';

export interface IPersonSpotlightCardProps {
  heading: string;
  /** Welcome Aboard and Promotions show this heading at different sizes — 20px and 18px respectively at the 1440px reference width. */
  headingFontSize: number;
  background: string;
  /** Hand-authored list (see data/mockData.ts) — the arrows cycle through these one at a time. */
  people: IPersonSpotlight[];
}

export interface IPersonSpotlightCardState {
  index: number;
  photoFailed: boolean;
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
    this.state = { index: 0, photoFailed: false };
  }

  private _prev = (): void => {
    this.setState(prev => ({ index: (prev.index - 1 + this.props.people.length) % this.props.people.length, photoFailed: false }));
  };

  private _next = (): void => {
    this.setState(prev => ({ index: (prev.index + 1) % this.props.people.length, photoFailed: false }));
  };

  private _onPhotoError = (): void => {
    this.setState({ photoFailed: true });
  };

  public render(): React.ReactElement {
    const { heading, headingFontSize, background, people } = this.props;
    const { photoFailed } = this.state;
    const person = people[this.state.index];
    if (!person) {
      // Author deleted every entry via the property pane's collection editor.
      return <></>;
    }

    const nameFontSize = fitFontSize(person.name, 18, 11, 16, 0.3);
    const titleFontSize = fitFontSize(person.title, 13, 9, 20, 0.15);
    const newTitleFontSize = person.newTitle ? fitFontSize(person.newTitle, 13, 9, 18, 0.15) : undefined;

    return (
      <div className={styles.card} style={{ background }}>
        {people.length > 1 && (
          <button className={styles.arrowButton} onClick={this._prev} aria-label="Previous">←</button>
        )}

        <p className={styles.heading} style={{ fontSize: headingFontSize }}>{heading}</p>

        {person.photoUrl && !photoFailed ? (
          <img className={styles.avatarPhoto} src={person.photoUrl} alt={person.name} onError={this._onPhotoError} />
        ) : (
          <span className={styles.avatar}>{person.initials}</span>
        )}
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
              <span className={styles.fieldValue} style={{ fontSize: fitFontSize(f.value, 12, 9, 12, 0.2) }}>
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
