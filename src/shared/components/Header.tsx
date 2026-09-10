import * as React from 'react';
import styles from './Header.module.scss';
import { NAV_ITEMS } from './navData';

export interface IHeaderProps {
  onOpenAssistant: () => void;
}

// The dropdown items are stand-ins for real SharePoint pages that don't
// exist yet — each gets its own distinct, stable slug so links behave like
// real (if currently empty) destinations rather than an inert "#" for
// every item, ready to be swapped for the actual page URL once it exists.
function slugify(label: string): string {
  return `#${label
    .toLowerCase()
    .replace(/&/g, 'and')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/(^-|-$)/g, '')}`;
}

export interface IHeaderState {
  openItem?: string;
}

export default class Header extends React.Component<IHeaderProps, IHeaderState> {
  private _closeTimeout?: number;

  constructor(props: IHeaderProps) {
    super(props);
    this.state = {};
  }

  public componentWillUnmount(): void {
    if (this._closeTimeout) {
      window.clearTimeout(this._closeTimeout);
    }
  }

  private _open = (label: string): void => {
    if (this._closeTimeout) {
      window.clearTimeout(this._closeTimeout);
      this._closeTimeout = undefined;
    }
    this.setState({ openItem: label });
  };

  // Closing on a short delay (instead of immediately) tolerates the brief
  // moment the cursor is between the nav link and the dropdown panel —
  // without it, fast or slightly-off-course mouse movement closes the menu
  // before the pointer ever reaches it.
  private _scheduleClose = (): void => {
    this._closeTimeout = window.setTimeout(() => {
      this.setState({ openItem: undefined });
    }, 200);
  };

  private _toggle = (label: string): void => {
    this.setState(prev => ({ openItem: prev.openItem === label ? undefined : label }));
  };

  public render(): React.ReactElement<IHeaderProps> {
    const { openItem } = this.state;

    return (
      <header className={styles.header}>
        <div className={styles.bar}>
          <div className={styles.brand}>
            <svg className={styles.brandMark} width="24" height="23" viewBox="0 0 24 23" fill="none" aria-hidden="true">
              <path d="M4 23H0V0H24L19.1111 9.11321H16.8889L12.8889 1.30189L9.33333 9.11321H1.33333L7.55556 15.6226L4 23Z" fill="#be002b" />
            </svg>
            <span className={styles.brandName}>Compass</span>
          </div>

          <ul className={styles.nav}>
            {NAV_ITEMS.map(item => (
              <li
                key={item.label}
                className={styles.navItem}
                onMouseEnter={() => item.columns && this._open(item.label)}
                onMouseLeave={() => item.columns && this._scheduleClose()}
              >
                <a
                  href="#"
                  className={`${styles.navLink} ${openItem === item.label ? styles.navLinkActive : ''}`}
                  onClick={e => {
                    e.preventDefault();
                    if (item.columns) {
                      this._toggle(item.label);
                    }
                  }}
                >
                  {item.label}
                </a>

                {item.columns && openItem === item.label && (
                  <div className={styles.dropdown}>
                    <div className={styles.dropdownInner}>
                      {item.columns.map((column, columnIndex) => (
                        <div className={styles.dropdownColumn} key={columnIndex}>
                          {column.map((group, groupIndex) => (
                            <div className={styles.dropdownGroup} key={groupIndex}>
                              {group.heading && <div className={styles.dropdownHeading}>{group.heading}</div>}
                              {group.items.map(link => (
                                <a
                                  key={link}
                                  href={slugify(link)}
                                  className={styles.dropdownLink}
                                >
                                  {link}
                                </a>
                              ))}
                            </div>
                          ))}
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </li>
            ))}
          </ul>

          <button type="button" className={styles.assistant} onClick={this.props.onOpenAssistant}>
            <span className={styles.assistantLabel}>Ask Your AI Assistant</span>
            <span className={styles.assistantSubmit} aria-hidden="true">→</span>
          </button>
        </div>
      </header>
    );
  }
}
