import * as React from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './FleetTable.module.scss';
import {
  FLEET_LIST_TITLE,
  ensureFleetList,
  formatDwt,
  loadVessels,
  sortVessels,
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
  /** Whether the signed-in user can create lists — only then is the "create the list" button offered. */
  canManageLists: boolean;
}

type Status = 'loading' | 'ready' | 'missing' | 'unreadable';

export interface IFleetTableState {
  status: Status;
  vessels: IVessel[];
  listUrl?: string;
  busy: boolean;
  error?: string;
}

function ShipIcon(): React.ReactElement {
  return (
    <svg className={styles.shipIcon} viewBox="0 0 66 18" aria-hidden="true" focusable="false">
      <g fill="currentColor">
        <path d="M0 9.5h66l-5.5 7.5H5.2z" />
        <rect x="47" y="2" width="9" height="7.5" />
        <rect x="49" y="0" width="5" height="2" />
        <rect x="6" y="6.2" width="7" height="3.3" />
        <rect x="15" y="6.2" width="7" height="3.3" />
        <rect x="24" y="6.2" width="7" height="3.3" />
        <rect x="33" y="6.2" width="7" height="3.3" />
        <rect x="9" y="3.4" width="1" height="2.8" />
        <rect x="27" y="3.4" width="1" height="2.8" />
      </g>
    </svg>
  );
}

export default class FleetTable extends React.Component<IFleetTableProps, IFleetTableState> {
  constructor(props: IFleetTableProps) {
    super(props);
    this.state = { status: 'loading', vessels: [], busy: false };
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
    this.setState({ status: 'ready', vessels: result.vessels });
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

  public render(): React.ReactElement {
    const { status, vessels, busy, error, listUrl } = this.state;
    const { sortBy, sortDescending, maxRows, canManageLists } = this.props;

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

    const sorted = sortVessels(vessels, sortBy, sortDescending);
    const rows = maxRows > 0 ? sorted.slice(0, maxRows) : sorted;

    return (
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
    );
  }
}
