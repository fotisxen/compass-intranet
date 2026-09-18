import * as React from 'react';
import * as ReactDom from 'react-dom';
import HelloWorld from '../src/webparts/helloWorld/components/HelloWorld';
import PublicHolidays from '../src/shared/components/PublicHolidays';
import EventsWidget from '../src/shared/components/EventsWidget';
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

// The stock-quote route on the same Function App needs no key/proxy (it's a
// plain public CORS-enabled endpoint), so unlike fleetApiUrl above this one
// defaults straight to the real production URL — override with
// ?stockApiUrl=... only if you need to point at something else.
const stockApiUrl =
  new URLSearchParams(window.location.search).get('stockApiUrl') ||
  'https://bpcstarbulkwebapi.azurewebsites.net/api/GetStockPrice';

// Proxies starbulk.com's own Euronext Athens feed (api/src/functions/euronextStock.ts)
// — not deployed anywhere yet, so unlike stockApiUrl above there's no known
// production URL to default to. Run `npm start` in api/ and open this
// preview as http://localhost:5600/?euronextApiUrl=http://localhost:7071/api/euronextStock
const euronextApiUrl = new URLSearchParams(window.location.search).get('euronextApiUrl') || undefined;

ReactDom.render(React.createElement(ChromeRoot, { chatApiUrl: '' }), document.getElementById('chrome-top'));

// Compass Holidays is its own standalone web part now, placed right under
// the nav — same original position, own container/row. No sizing/alignment
// here: PublicHolidays' own .zone/.outer/.widget layers already span the
// full row and align their content to the right, same as a real SharePoint
// zone div (a plain block, same as #chrome-bottom below for Footer).
ReactDom.render(
  React.createElement(
    'div',
    { style: { paddingTop: 24, background: '#ffffff' } },
    React.createElement(PublicHolidays, { spHttpClient: fakeSpHttpClient as never, siteUrl: 'https://example.sharepoint.com/sites/demo', offsetX: 50 })
  ),
  document.getElementById('holidays-root')
);

// Compass Events and Compass Fleet Map are both standalone web parts —
// this reproduces dropping them side by side into a real SharePoint
// 2-column section (same flex row/gap/stretch the original heroRow used),
// which is what actually produces the matched-height layout.
ReactDom.render(
  React.createElement(
    'div',
    {
      style: {
        display: 'flex',
        gap: 24,
        alignItems: 'stretch',
        maxWidth: 1440,
        margin: '0 auto',
        // No left/right padding — matches .main's own actual (disabled)
        // padding, so this lines up the same way real SharePoint zones do.
        padding: '20px 0 0',
        boxSizing: 'border-box',
        background: '#ffffff'
      }
    },
    React.createElement(EventsWidget, { spHttpClient: fakeSpHttpClient as never, siteUrl: 'https://example.sharepoint.com/sites/demo' }),
    React.createElement(
      'div',
      { style: { flex: 1, minWidth: 0, display: 'flex', flexDirection: 'column' } },
      React.createElement(FleetMap, { fleetApiUrl })
    )
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
ReactDom.render(React.createElement(Footer, { companyName: 'Compass', stockApiUrl, euronextApiUrl, contentOffsetX: 100 }), document.getElementById('chrome-bottom'));
