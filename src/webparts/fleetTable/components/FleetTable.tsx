import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './FleetTable.module.scss';
import classShipImage from '../../../shared/components/assets/fleet/ship-class.png';
import {
  FLEET_LIST_TITLE,
  ensureFleetList,
  fleetSummary,
  formatDwt,
  loadVessels,
  sortVessels,
  vesselClasses,
  type IVessel,
  type VesselSortField
} from '../../../shared/components/fleetVesselsList';

export interface IFleetTableProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  listTitle?: string;
  sortBy: VesselSortField;
  sortDescending: boolean;
  /** 0 = show every vessel. */
  maxRows: number;
  /** The row of filter buttons (All Types + one per Class). */
  showFilters: boolean;
  /** The "9 TANKERS RANGING FROM … DWT" line above the rows. */
  showSummary: boolean;
  /** The row of ship pictures with "17 x Kamsarmax" under them, above the filter buttons. */
  showClassCards: boolean;
  /** Plural name used in the summary line, e.g. TANKERS. */
  summaryNoun: string;
  /** Whether the signed-in user can create lists — only then is the "create the list" button offered. */
  canManageLists: boolean;
}

type Status = 'loading' | 'ready' | 'missing' | 'unreadable';

export interface IFleetTableState {
  status: Status;
  vessels: IVessel[];
  /** Whether the list has the Class column the filter buttons need. */
  hasClassColumn: boolean;
  /** The class chosen with the filter buttons; empty = All Types. */
  selectedClass: string;
  listUrl?: string;
  busy: boolean;
  error?: string;
}

function ShipIcon(): React.ReactElement {
  return <img className={styles.shipIcon} src={classShipImage} alt="" />;
}

export default class FleetTable extends React.Component<IFleetTableProps, IFleetTableState> {
  constructor(props: IFleetTableProps) {
    super(props);
    this.state = { status: 'loading', vessels: [], hasClassColumn: true, selectedClass: '', busy: false };
  }

  public componentDidMount(): void {
    this._load().catch(() => this.setState({ status: 'unreadable' }));
  }

  public componentDidUpdate(prev: IFleetTableProps): void {
    if (prev.listTitle !== this.props.listTitle) {
      this.setState({ status: 'loading' });
      this._load().catch(() => this.setState({ status: 'unreadable' }));
    }
  }

  private get _listTitle(): string {
    return (this.props.listTitle || '').trim() || FLEET_LIST_TITLE;
  }

  private async _load(): Promise<void> {
    const { spHttpClient, siteUrl } = this.props;
    const result = await loadVessels(spHttpClient, siteUrl, this._listTitle);
    if (result === 'missing' || result === 'unreadable') {
      this.setState({ status: result });
      return;
    }
    this.setState({ status: 'ready', vessels: result.vessels, hasClassColumn: result.hasClassColumn, selectedClass: '' });
  }

  private _createList = (): void => {
    const { spHttpClient, siteUrl } = this.props;
    this.setState({ busy: true, error: undefined });
    ensureFleetList(spHttpClient, siteUrl, this._listTitle)
      .then(list => {
        this.setState({ busy: false, listUrl: list.DefaultViewUrl });
        return this._load();
      })
      .catch((err: Error) => this.setState({ busy: false, error: err.message }));
  };

  private _addClassColumn = (): void => {
    const { spHttpClient, siteUrl } = this.props;
    this.setState({ busy: true, error: undefined });
    ensureFleetList(spHttpClient, siteUrl, this._listTitle)
      .then(() => {
        this.setState({ busy: false });
        return this._load();
      })
      .catch((err: Error) => this.setState({ busy: false, error: err.message }));
  };

  public render(): React.ReactElement {
    const { status, vessels, busy, error, listUrl, hasClassColumn, selectedClass } = this.state;
    const { sortBy, sortDescending, maxRows, canManageLists, showFilters, showSummary, showClassCards, summaryNoun } = this.props;

    if (status === 'loading') {
      return <div className={styles.table} aria-busy="true" />;
    }

    // Visitors never see setup messages — an unset-up list just renders nothing for them.
    if (status === 'missing' || status === 'unreadable') {
      if (!canManageLists) {
        return <></>;
      }
      return (
        <div className={styles.notice}>
          <p className={styles.noticeTitle}>
            {status === 'missing'
              ? `The list "${this._listTitle}" doesn't exist yet.`
              : `The list "${this._listTitle}" exists but is missing some columns.`}
          </p>
          <p className={styles.noticeText}>
            Create it here (owners only) and fill it in with one row per vessel — the table updates by itself.
          </p>
          <button type="button" className={styles.noticeButton} onClick={this._createList} disabled={busy}>
            {busy ? 'Setting up…' : status === 'missing' ? 'Create the list' : 'Add the missing columns'}
          </button>
          {error && <p className={styles.noticeError}>{error}</p>}
        </div>
      );
    }

    if (vessels.length === 0) {
      return canManageLists ? (
        <div className={styles.notice}>
          <p className={styles.noticeTitle}>The list &quot;{this._listTitle}&quot; has no vessels yet.</p>
          <p className={styles.noticeText}>
            Add rows to it (Title = vessel name, plus type, DWT, built year and shipyard) and they appear here.
            {listUrl && <> <a href={listUrl}>Open the list</a></>}
          </p>
        </div>
      ) : <></>;
    }

    // The filter buttons, then the summary and the rows of whatever they select.
    const allClasses = vesselClasses(vessels);
    const classes = showFilters ? allClasses : [];
    const columns = Math.max(allClasses.length + 1, 2);
    const counts: Record<string, number> = {};
    vessels.forEach(v => {
      if (v.vesselClass) {
        counts[v.vesselClass] = (counts[v.vesselClass] || 0) + 1;
      }
    });
    const active = classes.indexOf(selectedClass) !== -1 ? selectedClass : '';
    const shown = active ? vessels.filter(v => v.vesselClass === active) : vessels;
    const sorted = sortVessels(shown, sortBy, sortDescending);
    const rows = maxRows > 0 ? sorted.slice(0, maxRows) : sorted;

    return (
      <div className={styles.wrap} style={{ '--fleet-columns': columns } as React.CSSProperties}>
        {showSummary && rows.length > 0 && <div className={styles.summary}>{fleetSummary(shown, summaryNoun)}</div>}
        {showClassCards && allClasses.length > 0 && (
          <div className={styles.classCards}>
            {allClasses.map(c => (
              <div className={styles.classCard} key={c}>
                <img className={styles.classShip} src={classShipImage} alt="" />
                <div className={styles.classCount}>{counts[c]} x {c}</div>
              </div>
            ))}
          </div>
        )}
        {classes.length > 0 && (
          <div className={styles.filters} role="group" aria-label="Filter by type">
            {['', ...classes].map(c => (
              <button
                key={c || 'all'}
                type="button"
                className={`${styles.filter} ${c === active ? styles.filterActive : ''}`}
                aria-pressed={c === active}
                onClick={() => this.setState({ selectedClass: c })}
              >
                {c || 'All Types'}
              </button>
            ))}
          </div>
        )}
        {showFilters && canManageLists && !hasClassColumn && (
          <div className={styles.ownerHint}>
            The fleet list has no &quot;Class&quot; column yet, so the filter buttons cannot show.{' '}
            <button type="button" className={styles.noticeButton} onClick={this._addClassColumn} disabled={busy}>{busy ? 'Adding…' : 'Add the Class column'}</button>
            {error && <span className={styles.noticeError}> {error}</span>}
          </div>
        )}
        {showFilters && canManageLists && hasClassColumn && classes.length === 0 && (
          <div className={styles.ownerHint}>Fill in the &quot;Class&quot; column of the vessels (e.g. MR, LR1) and the filter buttons appear here.</div>
        )}
      <div className={styles.table} role="table" aria-label="Fleet">
        {rows.map(v => (
          <div className={styles.row} role="row" key={v.id}>
            <div className={styles.cellIcon} role="cell"><ShipIcon /></div>
            <div className={styles.cellName} role="cell">{v.name}</div>
            <div className={styles.cellMuted} role="cell">{v.type}</div>
            <div className={styles.cell} role="cell">{v.dwt !== undefined ? `DWT: ${formatDwt(v.dwt)}` : ''}</div>
            <div className={styles.cell} role="cell">{v.built ? `Built: ${v.built}` : ''}</div>
            <div className={styles.cell} role="cell">{v.shipyard ? `Shipyard: ${v.shipyard}` : ''}</div>
          </div>
        ))}
      </div>
      </div>
    );
  }
}
