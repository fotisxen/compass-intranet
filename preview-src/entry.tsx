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

// Compass Fleet Map is a separate web part from Compass Home now, so on the
// real page the closest match to "map beside Upcoming Events" is a 2-column
// section: Compass Home in the left column, Compass Fleet Map in the right
// — since Compass Home starts with Holidays then Events, the map ends up
// sitting visually next to that top portion, not scoped to Events alone.
// Mirrored here as one flex row instead of two separate containers.
ReactDom.render(
  React.createElement(
    'div',
    {
      style: {
        display: 'flex',
        alignItems: 'flex-start',
        gap: 24,
        maxWidth: 1440,
        margin: '0 auto',
        boxSizing: 'border-box',
        background: '#ffffff'
      }
    },
    // .main inside HelloWorld already contributes the page's standard 64px
    // left margin via its own padding, so this column adds none itself.
    React.createElement(
      'div',
      { style: { flex: 1, minWidth: 0 } },
      React.createElement(HelloWorld, {
        spHttpClient: fakeSpHttpClient as never,
        siteUrl: 'https://example.sharepoint.com/sites/demo'
      })
    ),
    // Padding-right here stands in for the page's standard 64px right
    // margin, since this column sits outside HelloWorld's own .main.
    React.createElement(
      'div',
      { style: { width: 500, flexShrink: 0, marginTop: 24, paddingRight: 64, boxSizing: 'border-box' } },
      React.createElement(FleetMap, { fleetApiUrl })
    )
  ),
  document.getElementById('root')
);

ReactDom.render(React.createElement(Footer, { companyName: 'Compass' }), document.getElementById('chrome-bottom'));
