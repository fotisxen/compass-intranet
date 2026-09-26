import * as React from 'react';
import * as ReactDom from 'react-dom';
import HelloWorld from '../src/webparts/helloWorld/components/HelloWorld';
import PublicHolidays from '../src/shared/components/PublicHolidays';
import EventsWidget from '../src/shared/components/EventsWidget';
import FleetMap from '../src/shared/components/FleetMap';
import ChromeRoot from '../src/extensions/compassChrome/components/ChromeRoot';
import Footer from '../src/shared/components/Footer';
import { welcomeAboard as sampleWelcome, promotion as samplePromotions } from '../src/webparts/helloWorld/components/data/mockData';

// No real SharePoint session here, so News/Holidays/Events' fetches are made
// to fail fast and fall back to their built-in mock data — same outcome a
// real site with those lists empty would produce. SharePoint REST can't
// realistically be hit from this standalone origin anyway (no auth cookie,
// no CORS allowance from the tenant) — see hosted workbench for real data.
const fakeSpHttpClient = {
  get: async () => ({ ok: false })
};

// Home's news carousel: 6 sample promoted pages, plus one pinned page that
// isn't promoted (fetched by path), so paging and pinning both show up.
const sampleNews = ['Fleet renewal programme reaches a new milestone', 'Summer safety campaign results', 'New offices open in Limassol', 'Quarterly results announced', 'Crew welfare initiative expands', 'Sustainability report published']
  .map((Title, i) => ({ Title, FileRef: '/sites/demo/SitePages/News-' + (i + 1) + '.aspx', Created: '2026-09-' + (20 - i) + 'T09:00:00Z', BannerImageUrl: null }));
const fakeHomeClient = {
  get: async (url: string) => {
    if (url.indexOf('GetFileByServerRelativePath') !== -1) {
      return { ok: true, json: async () => ({ Title: 'PINNED: Welcome to the new intranet', FileRef: '/sites/demo/SitePages/My-Pinned.aspx', Created: '2026-08-01T09:00:00Z', BannerImageUrl: null }) };
    }
    if (url.indexOf("GetByTitle('Site Pages')") !== -1) {
      return { ok: true, json: async () => ({ value: sampleNews }) };
    }
    return { ok: false };
  }
};

// Footer's Public Holidays column gets two sample rows (the client above fails on purpose).
const fakeHolidaysClient = {
  get: async () => ({
    ok: true,
    json: async () => ({ value: [{ Title: 'GR offices closed', EventDate: '2026-10-28T00:00:00Z' }, { Title: 'Cyprus Independence Day - CY', EventDate: '2026-11-02T00:00:00Z' }] })
  })
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
const euronextApiUrl = new URLSearchParams(window.location.search).get('euronextApiUrl') || 'https://mock-euronext.local/api/euronextStock';

// bpcstarbulkwebapi doesn't allow cross-origin calls from localhost, so
// Open Positions gets a small sample of the real GetJobListings shape here.
const realFetch = window.fetch.bind(window);
window.fetch = (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
  if (String(input).indexOf('mock-fleet') !== -1) {
    return Promise.resolve(new Response(JSON.stringify([{ name: 'Star Test I', lat: 37.9, lng: 23.7, status: 'At port' }, { name: 'Star Test II', lat: 36.2, lng: 26.5, status: 'Underway using engine', heading: 90 }, { name: 'Star Test III', lat: 40.1, lng: 25.0, status: 'Underway using engine', heading: 200 }]), { status: 200 }));
  }
  if (String(input).indexOf('GetStockPrice') !== -1) {
    return Promise.resolve(new Response(JSON.stringify({ data: [{ symbol: 'SBLK', lastTrade: 32.11, changePercent: 1.1, isDefault: 'true' }] }), { status: 200 }));
  }
  if (String(input).indexOf('mock-euronext') !== -1) {
    return Promise.resolve(new Response(JSON.stringify({ lastTrade: '28.13', changePercent: '2.2166' }), { status: 200 }));
  }
  if (String(input).indexOf('GetJobListings') !== -1) {
    const sample = [
      { title: 'Technical Coordinator', department: 'Technical' },
      { title: 'Fleet Manager', department: 'Technical' },
      { title: 'Junior Financial Reporting Officer', department: 'Financial Reporting Department' },
      { title: 'Summer internship', department: '' },
      { title: 'IT Officer', department: 'IT' }
    ];
    return Promise.resolve(new Response(JSON.stringify(sample), { status: 200 }));
  }
  return realFetch(input, init);
};

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
      React.createElement(FleetMap, { fleetApiUrl, englishLabels: new URLSearchParams(window.location.search).get('englishMap') === '1' })
    )
  ),
  document.getElementById('fleet-map-root')
);

ReactDom.render(
  React.createElement(HelloWorld, {
    spHttpClient: fakeHomeClient as never,
    siteUrl: 'https://example.sharepoint.com/sites/demo',
    pinnedNews: ['My-Pinned.aspx'],
    // Two entries each so the cards show their prev/next arrows.
    welcomeAboard: [...sampleWelcome, ...sampleWelcome],
    promotions: [...samplePromotions, ...samplePromotions]
  }),
  document.getElementById('root')
);
ReactDom.render(React.createElement(Footer, { companyName: 'Compass', stockApiUrl, euronextApiUrl, contentOffsetX: 100, spHttpClient: fakeHolidaysClient as never, siteUrl: 'https://example.sharepoint.com/sites/demo' }), document.getElementById('chrome-bottom'));
