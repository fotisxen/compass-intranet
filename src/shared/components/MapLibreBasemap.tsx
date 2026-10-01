import * as React from 'react';
import { useMap } from 'react-leaflet';
import type * as L from 'leaflet';
import { maplibreGL } from '@maplibre/maplibre-gl-leaflet';

// OpenFreeMap: free, no API key, OpenStreetMap data. Unlike raster tiles (which
// have each country's local script baked into the picture), vector tiles keep
// the names as data, so the label language can be chosen. Their terms require
// the attribution below to stay visible.
const STYLE_URL = 'https://tiles.openfreemap.org/styles/liberty';
const ATTRIBUTION = '<a href="https://openfreemap.org" target="_blank" rel="noopener noreferrer">OpenFreeMap</a> &copy; <a href="https://openmaptiles.org/" target="_blank" rel="noopener noreferrer">OpenMapTiles</a> Data from <a href="https://www.openstreetmap.org/copyright" target="_blank" rel="noopener noreferrer">OpenStreetMap</a>';
const MAPLIBRE_CSS_URL = 'https://unpkg.com/maplibre-gl@5/dist/maplibre-gl.css';
const MAPLIBRE_CSS_ID = 'maplibre-cdn-css';

// Same reason Leaflet's own CSS comes from the CDN in FleetMap.tsx.
function ensureMaplibreCss(): void {
  if (document.getElementById(MAPLIBRE_CSS_ID)) {
    return;
  }
  const link = document.createElement('link');
  link.id = MAPLIBRE_CSS_ID;
  link.rel = 'stylesheet';
  link.href = MAPLIBRE_CSS_URL;
  document.head.appendChild(link);
}

interface IStyleLayer {
  id: string;
  type: string;
  layout?: Record<string, unknown>;
  paint?: Record<string, unknown>;
}

interface IStyleJson {
  layers: IStyleLayer[];
}

// The stock style labels places as "Latin name + local-script name" (e.g.
// "Athens Αθήνα"); this keeps only the English name, falling back to the
// Latin-script one and then whatever the map has.
const ENGLISH_ONLY = ['coalesce', ['get', 'name_en'], ['get', 'name:latin'], ['get', 'name']];

// Flat cream for every land layer (same shade as the background, so the
// wood/grass/sand/residential patches that used to tint parts of the land
// green/tan/grey all disappear into one uniform color) plus the original
// OSM-like water and road colors — everything else about the style (roads,
// water, labels) is unchanged.
const LAND_COLOR = '#f2efe9';
const OSM_PAINT: Record<string, Record<string, string | number>> = {
  background: { 'background-color': LAND_COLOR },
  water: { 'fill-color': '#aad3df' },
  waterway_river: { 'line-color': '#aad3df' },
  waterway_other: { 'line-color': '#aad3df' },
  waterway_tunnel: { 'line-color': '#aad3df' },
  // park also has its own separate green outline color (fill-outline-color)
  // besides the fill itself — both need covering, plus the park_outline line
  // layer drawn on top of it.
  park: { 'fill-color': LAND_COLOR, 'fill-outline-color': LAND_COLOR },
  park_outline: { 'line-color': LAND_COLOR },
  landcover_wood: { 'fill-color': LAND_COLOR },
  landcover_grass: { 'fill-color': LAND_COLOR },
  landcover_sand: { 'fill-color': LAND_COLOR },
  landcover_ice: { 'fill-color': LAND_COLOR },
  // Pattern-filled (not a plain color), so it's hidden outright instead.
  landcover_wetland: { 'fill-opacity': 0 },
  landuse_residential: { 'fill-color': LAND_COLOR },
  landuse_cemetery: { 'fill-color': LAND_COLOR },
  landuse_school: { 'fill-color': LAND_COLOR },
  landuse_hospital: { 'fill-color': LAND_COLOR },
  landuse_pitch: { 'fill-color': LAND_COLOR },
  landuse_track: { 'fill-color': LAND_COLOR },
  aeroway_fill: { 'fill-color': LAND_COLOR },
  road_motorway: { 'line-color': '#e892a2' },
  road_trunk_primary: { 'line-color': '#f9b29c' },
  road_secondary_tertiary: { 'line-color': '#fdd7a1' }
};

function englishOnly(style: IStyleJson): IStyleJson {
  // The shaded-relief photo drawn under everything at low zoom is what makes
  // the stock style look so much more colourful; land is flat like OSM instead.
  // Country/region border lines are dropped too (place names stay).
  style.layers = style.layers.filter(layer => layer.id !== 'natural_earth' && layer.id.indexOf('boundary') !== 0);
  style.layers.forEach(layer => {
    const paint = OSM_PAINT[layer.id];
    if (paint) {
      layer.paint = { ...(layer.paint || {}), ...paint };
    }
    if (layer.type === 'symbol' && layer.layout && Array.isArray(layer.layout['text-field'])
      && JSON.stringify(layer.layout['text-field']).indexOf('name') >= 0) {
      layer.layout['text-field'] = ENGLISH_ONLY;
    }
  });
  return style;
}

/** Drop-in replacement for the OSM TileLayer inside a react-leaflet MapContainer, with every label in English. */
export default function MapLibreBasemap(): React.ReactElement {
  const map = useMap();

  React.useEffect(() => {
    let cancelled = false;
    let layer: L.MaplibreGL | undefined;
    ensureMaplibreCss();
    map.attributionControl.addAttribution(ATTRIBUTION);

    fetch(STYLE_URL)
      .then(response => response.json())
      .then((style: IStyleJson) => {
        if (cancelled) {
          return;
        }
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        layer = maplibreGL({ style: englishOnly(style) as any });
        layer.addTo(map);
      })
      .catch(() => {
        // Style couldn't be loaded — the map area stays blank (markers still work).
      });

    return () => {
      cancelled = true;
      map.attributionControl.removeAttribution(ATTRIBUTION);
      if (layer) {
        map.removeLayer(layer);
      }
    };
  }, [map]);

  return <></>;
}
