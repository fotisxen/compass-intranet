import * as React from 'react';
import * as L from 'leaflet';
import { MapContainer, TileLayer, Marker, Popup } from 'react-leaflet';
import styles from './FleetMap.module.scss';
import { offices } from './data/fleetMockData';

const LEAFLET_CSS_URL = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css';
const LEAFLET_CSS_ID = 'leaflet-cdn-css';

// Leaflet's stylesheet and default marker/control images reference each
// other with bare relative paths that SPFx's CSS loader can't resolve when
// bundled locally, so both load from the CDN instead (the basemap tiles
// already require network access at runtime).
function ensureLeafletCss(): void {
  if (document.getElementById(LEAFLET_CSS_ID)) {
    return;
  }
  const link = document.createElement('link');
  link.id = LEAFLET_CSS_ID;
  link.rel = 'stylesheet';
  link.href = LEAFLET_CSS_URL;
  document.head.appendChild(link);
}

// Offices have no "underway"/"at port" state, so they keep Leaflet's
// default pin rather than the vessel-status markers below.
const officeIcon = L.icon({
  iconUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png',
  iconRetinaUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon-2x.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png',
  iconSize: [25, 41],
  iconAnchor: [12, 41],
  popupAnchor: [1, -34],
  shadowSize: [41, 41]
});

// Matches the production AskStarBulk fleet map: a plain circle for vessels
// sitting in port, and a heading-rotated arrow for vessels underway.
const PORT_ICON = L.divIcon({
  className: styles.portMarker,
  html:
    '<svg width="14" height="14" viewBox="0 0 14 14"><circle cx="7" cy="7" r="5.2" fill="#F2A93B" stroke="#FFFFFF" stroke-width="2"/></svg>',
  iconSize: [14, 14],
  iconAnchor: [7, 7]
});

function underwayIcon(headingDeg: number): L.DivIcon {
  return L.divIcon({
    className: styles.underwayMarker,
    html: `<svg width="18" height="18" viewBox="0 0 18 18" style="transform: rotate(${headingDeg}deg)"><polygon points="9,1 15,15 9,11 3,15" fill="#3C9F52" stroke="#FFFFFF" stroke-width="1.6" stroke-linejoin="round"/></svg>`,
    iconSize: [18, 18],
    iconAnchor: [9, 9]
  });
}

// Common maritime flag-of-convenience codes used across the fleet; falls
// back to the raw code for anything not listed rather than guessing.
const FLAG_NAMES: Record<string, string> = {
  LR: 'Liberia',
  MH: 'Marshall Islands',
  MT: 'Malta',
  PA: 'Panama',
  GR: 'Greece',
  SG: 'Singapore',
  BS: 'Bahamas',
  CY: 'Cyprus'
};

function flagName(code?: string): string | undefined {
  if (!code) {
    return undefined;
  }
  return FLAG_NAMES[code.toUpperCase()] || code;
}

function isAtPort(status?: string): boolean {
  return !!status && status.toLowerCase().indexOf('port') !== -1;
}

function statusLabel(status?: string): string {
  if (!status) {
    return 'Unknown';
  }
  return isAtPort(status) ? 'At Port' : 'At sea';
}

function pad2(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

// Renders as "11 Sep, 10:38" — the caller appends a unit suffix (" UTC")
// where the field calls for one.
function formatUtc(iso?: string): string | undefined {
  if (!iso) {
    return undefined;
  }
  const date = new Date(iso);
  if (isNaN(date.getTime())) {
    return undefined;
  }
  const day = date.getUTCDate();
  const month = date.toLocaleString('en-US', { month: 'short', timeZone: 'UTC' });
  const hh = pad2(date.getUTCHours());
  const mm = pad2(date.getUTCMinutes());
  return `${day} ${month}, ${hh}:${mm}`;
}

export interface IFleetMapProps {
  /** URL of the fleet-positions Azure Function proxy (see api/src/functions/fleet.ts). Omit to keep showing office pins. */
  fleetApiUrl?: string;
}

interface IVesselPosition {
  name: string;
  lat: number;
  lng: number;
  imo?: string;
  flag?: string;
  status?: string;
  shipType?: string;
  vesselSize?: string;
  sog?: number;
  cog?: number;
  heading?: number;
  destination?: string;
  eta?: string;
  posDt?: string;
}

export interface IFleetMapState {
  showFleet: boolean;
  vessels: IVesselPosition[];
  hasLiveData: boolean;
  isLoadingVessels: boolean;
}

// The real provider's response has been confirmed against a live call
// (see api/src/functions/fleet.ts) — this still reads a couple of likely
// field-name variants defensively, but the primary names below
// (vesselName/latitude/longitude/...) are the ones actually returned.
function normalizeVessel(raw: Record<string, unknown>): IVesselPosition | undefined {
  const lat = Number(raw.lat ?? raw.Lat ?? raw.latitude ?? raw.Latitude);
  const lng = Number(raw.lng ?? raw.lon ?? raw.Lon ?? raw.longitude ?? raw.Longitude);
  if (!isFinite(lat) || !isFinite(lng)) {
    return undefined;
  }
  const name = String(
    raw.name ?? raw.Name ?? raw.vesselName ?? raw.VesselName ?? raw.shipName ?? raw.ShipName ?? 'Vessel'
  );

  const imo = raw.imo ?? raw.IMO ?? raw.Imo;
  const sog = raw.sog ?? raw.speed ?? raw.Sog;
  const cog = raw.cog ?? raw.course ?? raw.Cog;
  const heading = raw.heading ?? raw.Heading ?? cog;

  return {
    name,
    lat,
    lng,
    imo: imo !== undefined ? String(imo) : undefined,
    flag: typeof raw.flag === 'string' ? raw.flag : undefined,
    status: typeof raw.status === 'string' ? raw.status : undefined,
    shipType: typeof raw.typeOfShip === 'string' ? raw.typeOfShip : undefined,
    vesselSize: typeof raw.vesselSize === 'string' ? raw.vesselSize : undefined,
    sog: sog !== undefined ? Number(sog) : undefined,
    cog: cog !== undefined ? Number(cog) : undefined,
    heading: heading !== undefined ? Number(heading) : undefined,
    destination: typeof raw.destination === 'string' ? raw.destination : undefined,
    eta: typeof raw.eta === 'string' ? raw.eta : undefined,
    posDt: typeof raw.posDt === 'string' ? raw.posDt : undefined
  };
}

export default class FleetMap extends React.Component<IFleetMapProps, IFleetMapState> {
  constructor(props: IFleetMapProps) {
    super(props);
    // On by default — the map opens already showing the fleet instead of
    // requiring a click on the "Our Fleet Live" tag first.
    this.state = { showFleet: true, vessels: [], hasLiveData: false, isLoadingVessels: false };
  }

  public componentDidMount(): void {
    ensureLeafletCss();
    if (this.state.showFleet && !this.state.hasLiveData) {
      this._startLoadingVessels();
    }
  }

  private _startLoadingVessels(): void {
    this.setState({ isLoadingVessels: true });
    this._loadVessels()
      .catch(() => {
        // Real positions unavailable (proxy not configured for this
        // environment, or the request failed) — the office pins already
        // in state stay as a fallback so the map never shows nothing.
      })
      .then(() => this.setState({ isLoadingVessels: false }))
      .catch(() => undefined);
  }

  private _toggleFleet = (): void => {
    const isTurningOn = !this.state.showFleet;
    this.setState(prev => ({ showFleet: !prev.showFleet }));

    if (isTurningOn && !this.state.hasLiveData) {
      this._startLoadingVessels();
    }
  };

  private async _loadVessels(): Promise<void> {
    const { fleetApiUrl } = this.props;
    if (!fleetApiUrl) {
      return;
    }

    // The live positions provider aggregates the whole fleet per request
    // and routinely takes 20+ seconds to respond — that's normal for this
    // endpoint, not a stalled request.
    const response = await fetch(fleetApiUrl);
    if (!response.ok) {
      return;
    }

    const data: unknown = await response.json();
    const rawList = Array.isArray(data) ? data : (data as { value?: unknown[] })?.value;
    if (!Array.isArray(rawList)) {
      return;
    }

    const vessels = rawList
      .map(item => normalizeVessel(item as Record<string, unknown>))
      .filter((v): v is IVesselPosition => v !== undefined);

    if (vessels.length > 0) {
      this.setState({ vessels, hasLiveData: true });
    }
  }

  private _iconFor(vessel: IVesselPosition): L.DivIcon {
    return isAtPort(vessel.status) ? PORT_ICON : underwayIcon(vessel.heading ?? vessel.cog ?? 0);
  }

  private _renderVesselCard(vessel: IVesselPosition): React.ReactElement {
    const position = formatUtc(vessel.posDt);
    const eta = formatUtc(vessel.eta);
    const speedCourse = [
      vessel.sog !== undefined ? `${vessel.sog} kn` : undefined,
      vessel.cog !== undefined ? `${vessel.cog}°` : undefined
    ]
      .filter(Boolean)
      .join(' · ');
    const idLine = [vessel.imo ? `IMO ${vessel.imo}` : undefined, flagName(vessel.flag)].filter(Boolean).join(' · ');
    const type = [vessel.shipType, vessel.vesselSize].filter(Boolean).join(' · ');

    return (
      <div className={styles.vesselCard}>
        <div className={styles.vesselCardHeader}>
          <span className={styles.vesselCardTitle}>{vessel.name}</span>
          <span className={isAtPort(vessel.status) ? styles.statusBadgeAtPort : styles.statusBadgeAtSea}>
            {statusLabel(vessel.status)}
          </span>
        </div>
        {idLine && <div className={styles.vesselCardSub}>{idLine}</div>}
        <div className={styles.vesselCardDetails}>
          {type && (
            <>
              <span className={styles.vesselCardLabel}>Type</span>
              <span className={styles.vesselCardValue}>{type}</span>
            </>
          )}
          {position && (
            <>
              <span className={styles.vesselCardLabel}>Position</span>
              <span className={styles.vesselCardValue}>{position} UTC</span>
            </>
          )}
          {speedCourse && (
            <>
              <span className={styles.vesselCardLabel}>Speed / course</span>
              <span className={styles.vesselCardValue}>{speedCourse}</span>
            </>
          )}
          {vessel.destination && (
            <>
              <span className={styles.vesselCardLabel}>Next port</span>
              <span className={styles.vesselCardValue}>{vessel.destination}</span>
            </>
          )}
          {eta && (
            <>
              <span className={styles.vesselCardLabel}>ETA</span>
              <span className={styles.vesselCardValue}>{eta}</span>
            </>
          )}
        </div>
      </div>
    );
  }

  public render(): React.ReactElement {
    const { showFleet, vessels, hasLiveData, isLoadingVessels } = this.state;

    return (
      <div className={styles.mapWrap}>
        <MapContainer
          className={styles.map}
          center={[25, 15]}
          zoom={3}
          zoomControl={false}
          scrollWheelZoom={false}
        >
          {/* Carto's free basemaps now require an API key for anonymous use
              (they watermark unauthenticated requests), so this uses standard
              OpenStreetMap tiles instead — genuinely free, no key needed.
              Their usage policy requires the attribution below to stay visible
              (it's just styled smaller/quieter via :global in the scss). */}
          <TileLayer
            attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
            url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
          />
          {showFleet && hasLiveData &&
            vessels.map(vessel => (
              <Marker
                key={`${vessel.imo ?? vessel.name}`}
                position={[vessel.lat, vessel.lng]}
                icon={this._iconFor(vessel)}
              >
                <Popup minWidth={220}>
                  {this._renderVesselCard(vessel)}
                </Popup>
              </Marker>
            ))}
          {showFleet && !hasLiveData && !isLoadingVessels &&
            offices.map(office => (
              <Marker key={office.city} position={[office.lat, office.lng]} icon={officeIcon}>
                <Popup>
                  <div className={styles.popup}>
                    <strong>{`${office.city}, ${office.country}`}</strong>
                    {office.address}
                  </div>
                </Popup>
              </Marker>
            ))}
        </MapContainer>
        {showFleet && isLoadingVessels && (
          <div className={styles.loadingOverlay}>
            <span className={styles.spinner} aria-hidden="true" />
            <span>Loading fleet positions…</span>
          </div>
        )}
        <button type="button" className={styles.tag} onClick={this._toggleFleet} disabled={isLoadingVessels}>
          {isLoadingVessels ? 'Loading fleet…' : 'Our Fleet Live'}
        </button>
      </div>
    );
  }
}
