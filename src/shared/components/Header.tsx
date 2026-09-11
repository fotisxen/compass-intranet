import * as React from 'react';
import styles from './Header.module.scss';
import { NAV_ITEMS } from './navData';

export interface IHeaderProps {
  onOpenAssistant: () => void;
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
                {item.columns ? (
                  <a
                    href="#"
                    className={`${styles.navLink} ${openItem === item.label ? styles.navLinkActive : ''}`}
                    onClick={e => {
                      e.preventDefault();
                      this._toggle(item.label);
                    }}
                  >
                    {item.label}
                  </a>
                ) : (
                  <a href={item.href || '#'} className={styles.navLink}>
                    {item.label}
                  </a>
                )}

                {item.columns && openItem === item.label && (
                  <div className={styles.dropdown}>
                    <div className={styles.dropdownInner}>
                      {item.columns.map((column, columnIndex) => (
                        <div className={styles.dropdownColumn} key={columnIndex}>
                          {column.map((group, groupIndex) => (
                            <div className={styles.dropdownGroup} key={groupIndex}>
                              {group.heading && (
                                group.headingHref ? (
                                  <a href={group.headingHref} className={styles.dropdownHeading}>
                                    {group.heading}
                                  </a>
                                ) : (
                                  <div className={styles.dropdownHeading}>{group.heading}</div>
                                )
                              )}
                              {group.items.map(link => (
                                <a
                                  key={link.label}
                                  href={link.href}
                                  className={styles.dropdownLink}
                                >
                                  {link.label}
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
