import * as React from 'react';
import * as ReactDom from 'react-dom';
import HomeNews from '../src/webparts/helloWorld/components/HomeNews';
import HomePeople from '../src/webparts/homePeople/components/HomePeople';
import PublicHolidays from '../src/shared/components/PublicHolidays';
import EventsWidget from '../src/shared/components/EventsWidget';
import FleetMap from '../src/shared/components/FleetMap';
import ChromeRoot from '../src/extensions/compassChrome/components/ChromeRoot';
import Footer from '../src/shared/components/Footer';
import { welcomeAboard as sampleWelcome, promotion as samplePromotions, anniversaries as sampleAnniversaries } from '../src/webparts/homePeople/components/data/mockData';

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
const samplePhoto = (i: number): string => 'data:image/svg+xml;utf8,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="90" height="90"><rect width="90" height="90" fill="' + (i ? '#c2410c' : '#0f766e') + '"/><circle cx="45" cy="36" r="17" fill="#fde68a"/><rect x="16" y="58" width="58" height="40" rx="20" fill="#fde68a"/></svg>');
const fakeHomeClient = {
  get: async (url: string) => {
    if (url.indexOf('GetFileByServerRelativePath') !== -1) {
      const m = /(Pinned-[0-9])/i.exec(decodeURIComponent(url));
      const name = m ? m[1] : 'My-Pinned';
      return { ok: true, json: async () => ({ Title: 'PINNED: ' + name, FileRef: '/sites/demo/SitePages/' + name + '.aspx', Created: '2026-08-01T09:00:00Z', BannerImageUrl: null }) };
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

import WideNewsCard from '../src/webparts/wideNewsCard/components/WideNewsCard';
import HighlightBanner from '../src/webparts/highlightBanner/components/HighlightBanner';
import FeatureSplit from '../src/webparts/featureSplit/components/FeatureSplit';
import ProfileLinkCard from '../src/webparts/profileLinkCard/components/ProfileLinkCard';
import TestimonialCarousel from '../src/webparts/testimonialCarousel/components/TestimonialCarousel';
import CircleIconList from '../src/webparts/circleIconList/components/CircleIconList';
import LogoGridFeature from '../src/webparts/logoGridFeature/components/LogoGridFeature';
import DualColumnFeature from '../src/webparts/dualColumnFeature/components/DualColumnFeature';
import StatFeature from '../src/webparts/statFeature/components/StatFeature';
import FleetTable from '../src/webparts/fleetTable/components/FleetTable';

function newRoot(): HTMLElement {
  const el = document.createElement('div');
  el.style.marginBottom = '32px';
  document.body.appendChild(el);
  return el;
}

ReactDom.render(React.createElement('div', { style: { padding: 24, border: '2px dashed #999' } }, React.createElement(WideNewsCard, {
  title: 'Fleet renewal programme reaches a new milestone',
  dateText: 'September 20, 2026',
  linkUrl: '#'
})), newRoot());

ReactDom.render(React.createElement(HighlightBanner, {
  title: '1. Eligibility', titleColor: '#ffffff',
  items: [
    'Coverage starts after your first 6 months of employment.',
    'Dependents can be added after 1 year of employment. Scroll down to the "Useful Resources" section to find the "Health Insurance Program Enrollment Form", complete it, and send it to HR through e-mail.'
  ],
  bodyColor: '#000000', bgColor: '#d9d9d9'
}), newRoot());

ReactDom.render(React.createElement(FeatureSplit, {
  title: 'Title', titleColor: '#082244',
  bodyText: 'Introduction of each Section, repeated a few times to show wrapping across the available width of this box.',
  bodyColor: '#082244', imageUrl: undefined
}), newRoot());

ReactDom.render(React.createElement(ProfileLinkCard, { topText: 'View My Profile', topLinkUrl: 'https://example.com/profile' }), newRoot());

ReactDom.render(React.createElement(TestimonialCarousel, {
  title: 'Title', titleColor: '#082244', bgColor: '#e6e6e6',
  testimonials: [1, 2, 3, 4, 5].map(i => ({ quote: 'When I donate blood I feel like I am giving life to someone in need — without expecting anything in return.', author: 'Dimitris Spyrou ' + i }))
}), newRoot());

ReactDom.render(React.createElement(CircleIconList, {
  circleText: '1', title: 'Text 1',
  items: ['List text 1', 'List text 2', 'List text 3']
}), newRoot());

ReactDom.render(React.createElement(LogoGridFeature, {
  imageUrl: undefined,
  logos: [1, 2, 3, 4, 5].map(i => ({ logoUrl: 'https://via.placeholder.com/60?text=' + i, linkUrl: 'https://example.com/' + i })),
  title: 'Title', bodyText: 'Call the Generali Call Center for 24/7 medical advice and phone guidance.'
}), newRoot());

ReactDom.render(React.createElement(DualColumnFeature, {
  bgColor: '#7ac0fb',
  leftTitle: 'Pre-hospitalization approval',
  leftItems: ['Option 1: Through MyGenerali app / portal.', 'Option 2: Through HR Department. Navigate to the "Useful Resources" section, download the forms, and send them to HR.'],
  rightTitle: 'Exclusions (high-level)',
  rightItems: ['Option 1: Through MyGenerali app / portal.', 'Option 2: Through HR Department. Navigate to the "Useful Resources" section, download the forms, and send them to HR.']
}), newRoot());

ReactDom.render(React.createElement(StatFeature, {
  bgColor: '#3ecf8e', eyebrowText: 'SECTION, CATEGORY, TOP OF EVERY PAGE', eyebrowColor: '#ffffff',
  titleText: 'Big Title', titleColor: '#ffffff',
  bodyText: 'Introduction of each Section, repeated a few times to show wrapping.', bodyColor: '#ffffff',
  imageUrl: undefined, columnsColor: '#ffffff',
  columns: [{ title: 'Title 1', text: 'Introduction of each section.' }, { title: 'Title 2', text: 'Introduction of each section.' }, { title: 'Title 3', text: 'Introduction of each section.' }]
}), newRoot());

// Fake SharePoint client for the list-driven pieces. With ?navList=1 the
// header reads a (made-up) Compass Navigation list that reorders, renames and
// adds entries, to see the dynamic menu react; without it, no client is passed
// and the header shows the built-in menu, as before.
const fleetSample = [
  ['Goliath', 'Newcastlemax', 209537, '2015', 'NACKS China'], ['Gargantua', 'Newcastlemax', 209529, '2015', 'NACKS China'],
  ['Star Gina 2GR', 'Newcastlemax', 209475, '2016', 'NACKS China'], ['Maharaj', 'Newcastlemax', 209472, '2015', 'NACKS China'],
  ['Star Leo', 'Newcastlemax', 207939, '2018', 'SWS China'], ['Star Laetitia', 'Newcastlemax', 207896, '2017', 'SWS, China'],
  ['Star Ariadne', 'Newcastlemax', 207812, '2017', 'SWS China'], ['Star Virgo', 'Newcastlemax', 207810, '2017', 'SWS China'],
  ['Star Libra', 'Newcastlemax', 207765, '2016', 'SWS China']
].map((v, i) => ({ Id: i + 1, Title: v[0], VesselType: v[1], Dwt: v[2], BuiltYear: v[3], Shipyard: v[4] }));
const navRows = [
  { Id: 1, Title: 'Fleet', NavUrl: '/fleet', SortOrder: 10, DropdownColumn: 1, ShowInMenu: true },
  { Id: 2, Title: 'Safety', NavUrl: '', SortOrder: 20, DropdownColumn: 1, ShowInMenu: true },
  { Id: 3, Title: 'Drills', NavUrl: '/drills', ParentItemId: 2, SortOrder: 10, DropdownColumn: 1, ShowInMenu: true },
  { Id: 4, Title: 'Emergency contacts', NavUrl: '/contacts', ParentItemId: 2, SortOrder: 10, DropdownColumn: 2, ShowInMenu: true },
  { Id: 5, Title: 'People Hub', NavUrl: '/people', SortOrder: 30, DropdownColumn: 1, ShowInMenu: true }
];
const fakeListClient = {
  get: async (url: string) => {
    if (url.indexOf('Compass%20Fleet') !== -1 || url.indexOf('Compass Fleet') !== -1) {
      return { ok: true, status: 200, json: async () => ({ value: fleetSample }) };
    }
    if (url.indexOf('Compass%20Navigation') !== -1 || url.indexOf('Compass Navigation') !== -1) {
      return { ok: true, status: 200, json: async () => ({ value: navRows }) };
    }
    return { ok: false, status: 404, json: async () => ({}) };
  }
};
const fakeEventsClient = {
  get: async () => ({
    ok: true,
    status: 200,
    json: async () => ({
      value: ['Safety drill at the Piraeus office', 'Annual company beach clean-up', 'Crew family day'].map((Title, i) => ({
        Id: i + 1, Title, EventDate: '2027-0' + (i + 1) + '-15T09:00:00Z', BannerUrl: null
      }))
    })
  })
};
const useFakeEvents = new URLSearchParams(window.location.search).get('events') === '1';
const useFakeNav = new URLSearchParams(window.location.search).get('navList') === '1';
ReactDom.render(React.createElement(FleetTable, { spHttpClient: fakeListClient as never, siteUrl: 'https://example.sharepoint.com/sites/demo', sortBy: 'dwt', sortDescending: true, maxRows: 0, canManageLists: true }), newRoot());

ReactDom.render(React.createElement(ChromeRoot, useFakeNav
  ? { chatApiUrl: '', spHttpClient: fakeListClient as never, siteUrl: 'https://example.sharepoint.com/sites/demo' }
  : { chatApiUrl: '' }), document.getElementById('chrome-top'));

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
    React.createElement(EventsWidget, { spHttpClient: (useFakeEvents ? fakeEventsClient : fakeSpHttpClient) as never, siteUrl: 'https://example.sharepoint.com/sites/demo' }),
    React.createElement(
      'div',
      { style: { flex: 1, minWidth: 0, display: 'flex', flexDirection: 'column' } },
      React.createElement(FleetMap, { fleetApiUrl, englishLabels: new URLSearchParams(window.location.search).get('englishMap') === '1' })
    )
  ),
  document.getElementById('fleet-map-root')
);

ReactDom.render(
  React.createElement(HomeNews, {
    spHttpClient: fakeHomeClient as never,
    siteUrl: 'https://example.sharepoint.com/sites/demo',
    // ?pinned=N (0-4) pins that many sample pages; default 1.
    pinnedNews: Array.from({ length: Number(new URLSearchParams(window.location.search).get('pinned') ?? 1) }, (_, i) => 'Pinned-' + (i + 1) + '.aspx')
  }),
  document.getElementById('root')
);
ReactDom.render(
  React.createElement(HomePeople, {
    // Two entries each so the cards show their prev/next arrows.
    // ?long=1 swaps in the longest realistic text, to check the cards keep one size.
    welcomeAboard: new URLSearchParams(window.location.search).get('long') === '1' ? [...sampleWelcome, { ...sampleWelcome[0], name: 'Konstantinos Papadimitriou-Vasileiou', title: 'Senior Technical Superintendent Fleet', fields: sampleWelcome[0].fields.map(f => ({ ...f, value: f.value + ' and Maritime Compliance Division Greece' })) }] : [...sampleWelcome, ...sampleWelcome],
    promotions: [...samplePromotions, ...samplePromotions],
    anniversaries: sampleAnniversaries.map((a, i) => (i < 2 ? { ...a, photoUrl: samplePhoto(i) } : a))
  }),
  document.getElementById('people-root')
);
ReactDom.render(React.createElement(Footer, { companyName: 'Compass', stockApiUrl, euronextApiUrl, contentOffsetX: 100, spHttpClient: fakeHolidaysClient as never, siteUrl: 'https://example.sharepoint.com/sites/demo' }), document.getElementById('chrome-bottom'));
