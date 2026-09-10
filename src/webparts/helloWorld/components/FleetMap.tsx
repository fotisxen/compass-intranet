import * as React from 'react';
import * as L from 'leaflet';
import { MapContainer, TileLayer, Marker, Popup } from 'react-leaflet';
import styles from './FleetMap.module.scss';
import { offices } from './data/mockData';

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

const officeIcon = L.icon({
  iconUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png',
  iconRetinaUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon-2x.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png',
  iconSize: [25, 41],
  iconAnchor: [12, 41],
  popupAnchor: [1, -34],
  shadowSize: [41, 41]
});

export interface IFleetMapState {
  showFleet: boolean;
}

export default class FleetMap extends React.Component<Record<string, never>, IFleetMapState> {
  constructor(props: Record<string, never>) {
    super(props);
    this.state = { showFleet: false };
  }

  public componentDidMount(): void {
    ensureLeafletCss();
  }

  private _toggleFleet = (): void => {
    this.setState(prev => ({ showFleet: !prev.showFleet }));
  };

  public render(): React.ReactElement {
    const { showFleet } = this.state;

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
          {showFleet && offices.map(office => (
            <Marker key={office.city} position={[office.lat, office.lng]} icon={officeIcon}>
              <Popup>
                <div className={styles.popup}>
                  <strong>{office.city}, {office.country}</strong>
                  {office.address}
                </div>
              </Popup>
            </Marker>
          ))}
        </MapContainer>
        <button type="button" className={styles.tag} onClick={this._toggleFleet}>
          Our Fleet Live
        </button>
      </div>
    );
  }
}
