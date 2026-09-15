import * as React from 'react';
import * as ReactDom from 'react-dom';
import HelloWorld from '../src/webparts/helloWorld/components/HelloWorld';
import FleetMap from '../src/shared/components/FleetMap';
import ChromeRoot from '../src/extensions/compassChrome/components/ChromeRoot';
import Footer from '../src/shared/components/Footer';

// No real SharePoint session here, so News/Holidays/Events' fetches are made
// to fail fast and fall back to their built-in mock data — same outcome a
// real site with those lists empty would produce. SharePoint REST can't
// realistically be hit from this standalone origin anyway (no auth cookie,
// no CORS allowance from the tenant) — see hosted workbench for real data.
const fakeSpHttpClient = {
  get: async () => ({ ok: false })
};

// The fleet-positions Azure Function is a normal CORS-enabled REST endpoint
// with no SharePoint session dependency, so — unlike the calls above — it
// genuinely can be hit from here. Run `npm start` in api/ (after filling in
// api/local.settings.json with your real FLEET_API_URL/FLEET_API_KEY) and
// open this preview as http://localhost:5600/?fleetApiUrl=http://localhost:7071/api/fleet
const fleetApiUrl = new URLSearchParams(window.location.search).get('fleetApiUrl') || undefined;

ReactDom.render(React.createElement(ChromeRoot, { chatApiUrl: '' }), document.getElementById('chrome-top'));

// Fleet Map is the only piece still split out of Compass Home — it's its
// own standalone web part (Compass Fleet Map), placed on the real page
// separately (above or below Compass Home, wherever you drop it).
ReactDom.render(
  React.createElement(
    'div',
    { style: { maxWidth: 1440, margin: '0 auto', padding: '24px 64px 0', boxSizing: 'border-box', background: '#ffffff' } },
    React.createElement(FleetMap, { fleetApiUrl })
  ),
  document.getElementById('fleet-map-root')
);

ReactDom.render(
  React.createElement(HelloWorld, {
    spHttpClient: fakeSpHttpClient as never,
    siteUrl: 'https://example.sharepoint.com/sites/demo'
  }),
  document.getElementById('root')
);
ReactDom.render(React.createElement(Footer, { companyName: 'Compass' }), document.getElementById('chrome-bottom'));
