import * as React from 'react';
import * as ReactDom from 'react-dom';
import HomeNews from '../src/webparts/helloWorld/components/HomeNews';
import NewsGrid from '../src/webparts/newsGrid/components/NewsGrid';
import MarketplaceFilters from '../src/webparts/marketplaceFilters/components/MarketplaceFilters';
import MarketplacePost from '../src/webparts/marketplacePost/components/MarketplacePost';
import MarketplaceInfo from '../src/webparts/marketplaceInfo/components/MarketplaceInfo';
import MarketplaceListings from '../src/webparts/marketplaceListings/components/MarketplaceListings';
import BackgroundBox from '../src/webparts/backgroundBox/components/BackgroundBox';
import SectionHeading from '../src/webparts/sectionHeading/components/SectionHeading';
import EventsGrid from '../src/webparts/eventsGrid/components/EventsGrid';
import EventImage from '../src/webparts/eventImage/components/EventImage';
import EventDetails from '../src/webparts/eventDetails/components/EventDetails';
import LinkButton from '../src/webparts/linkButton/components/LinkButton';
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
const sampleNews = ['Fleet renewal programme reaches a new milestone', 'Summer safety campaign results', 'New offices open in Limassol', 'Quarterly results announced', 'Crew welfare initiative expands', 'Sustainability report published', 'Port state control inspection results', 'New hires join the technical team', 'Charity run raises record amount', 'Fleet safety award received']
  .map((Title, i) => ({ Title, FileRef: '/sites/demo/SitePages/News-' + (i + 1) + '.aspx', Created: '2026-09-' + (20 - i) + 'T09:00:00Z', BannerImageUrl: null }));

// ---- Marketplace: an in-memory stand-in for the two SharePoint lists.
const mpPhoto = (c1: string, c2: string, label: string): string => 'data:image/svg+xml;utf8,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="640" height="400"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="' + c1 + '"/><stop offset="1" stop-color="' + c2 + '"/></linearGradient></defs><rect width="640" height="400" fill="url(#g)"/><text x="320" y="215" font-size="40" text-anchor="middle" fill="#ffffff" font-family="Arial">' + label + '</text></svg>');
const mpDay = (n: number): string => new Date(Date.now() - n * 86400000).toISOString();
const mpCategories = ['For sale', 'Free', 'Wanted', 'Swap', 'Housing', 'Carpool', 'Lost & found'];
const mpRow = (id: number, title: string, price: string, area: string, cat: string, days: number, author: string, photos: number, status = 0, details = 'Good condition, pick up from the office. Message me on Teams if interested.') => ({
  Id: id, Title: title, Price: price, Area: area, Details: details, ListingCategory: cat, Created: mpDay(days), OData__ModerationStatus: status,
  Author: { Title: author, EMail: author.toLowerCase().replace(/\s+/g, '.') + '@example.com' },
  AttachmentFiles: Array.from({ length: photos }, (_, i) => ({ FileName: 'p' + i + '.jpg', ServerRelativeUrl: mpPhoto(['#1d4d9a', '#0f6b63', '#9b2c52', '#8a5a00'][(id + i) % 4], '#082244', title.slice(0, 18) + ' ' + (i + 1)) }))
});
const mpListings = [
  mpRow(1, 'Baby car seat, 0–4 years', 'Free', 'Athens HQ', 'Free', 3, 'Giorgos Tsimpoukelis', 3),
  mpRow(2, 'IKEA desk lamp', '€8', 'Piraeus', 'For sale', 4, 'Maria Papadopoulou', 1),
  mpRow(3, "Kids' bike, 16 inch", '€40 or best offer', 'Athens HQ', 'For sale', 6, 'Nikos Andreou', 2),
  mpRow(4, 'Looking for a 2-bedroom flat near the office', 'Up to €700/month', 'Kifissia', 'Wanted', 7, 'Eleni Vasiliou', 0, 0, 'Moving to Athens in January. Happy to share, no pets.'),
  mpRow(5, 'Carpool Glyfada to Marousi, mornings', '€3 per ride', 'Glyfada', 'Carpool', 9, 'Dimitris Perivolakis', 0),
  mpRow(6, 'Swap: PlayStation 4 for Nintendo Switch', 'Swap', 'Athens HQ', 'Swap', 12, 'Anna Georgiou', 4),
  mpRow(7, 'Studio to rent, Piraeus port', '€450/month', 'Piraeus', 'Housing', 14, 'Giorgos Stefanidis', 3),
  mpRow(8, 'Found: black umbrella at reception', 'Free', 'Athens HQ', 'Lost & found', 15, 'Reception Desk', 1),
  mpRow(9, 'Standing desk, adjustable', '€120', 'Athens HQ', 'For sale', 18, 'Maria Dirla', 2),
  mpRow(10, 'Board games bundle (6 games)', 'Free', 'Kifissia', 'Free', 21, 'Nikos Kallinikos', 1)
];
const mpMine = [mpRow(11, 'Winter tyres 205/55 R16', '€150', 'Athens HQ', 'For sale', 1, 'Fotis Xenos', 1, 2), mpRow(12, 'IKEA desk lamp', '€8', 'Piraeus', 'For sale', 4, 'Fotis Xenos', 1, 0), mpRow(13, 'Old monitor', 'Free', 'Athens HQ', 'Free', 5, 'Fotis Xenos', 0, 1)];
const fakeMarketplaceClient = {
  get: async (url: string) => {
    const u = decodeURIComponent(url);
    if (u.indexOf('currentuser') !== -1) return { ok: true, status: 200, json: async () => ({ Id: 7 }) };
    if (u.indexOf("Compass Marketplace Categories')/items") !== -1) return { ok: true, status: 200, json: async () => ({ value: mpCategories.map((c, i) => ({ Id: i + 1, Title: c, SortOrder: i + 1 })) }) };
    if (u.indexOf("Compass Marketplace')/items") !== -1) return { ok: true, status: 200, json: async () => ({ value: u.indexOf('AuthorId') !== -1 ? mpMine : mpListings }) };
    if (u.indexOf("Compass Marketplace')") !== -1) return { ok: true, status: 200, json: async () => ({ Id: 'x', ListItemEntityTypeFullName: 'SP.Data.CompassMarketplaceListItem', DefaultViewUrl: '/sites/demo/Lists/Compass Marketplace/AllItems.aspx' }) };
    return { ok: false, status: 404 };
  },
  post: async (url: string) => {
    if (url.indexOf('AttachmentFiles/add') !== -1) return { ok: true, status: 200, json: async () => ({}) };
    return { ok: true, status: 201, json: async () => ({ Id: 99 }) };
  }
};

// ---- Events: an in-memory stand-in for the site's Events (calendar) list.
const evPhoto = (c1: string, c2: string): string => 'data:image/svg+xml;utf8,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="800" height="450"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="' + c1 + '"/><stop offset="1" stop-color="' + c2 + '"/></linearGradient></defs><rect width="800" height="450" fill="url(#g)"/><circle cx="560" cy="120" r="46" fill="#ffffff" opacity=".35"/><rect x="0" y="330" width="800" height="120" fill="#000000" opacity=".25"/></svg>');
const evRow = (id: number, title: string, startIso: string, endIso: string, allDay: boolean, category: string, location: string, withImage: boolean, about: string) => ({
  Id: id, Title: title, EventDate: startIso, EndDate: endIso, fAllDayEvent: allDay, Category: category, Location: location, Description: about,
  BannerUrl: withImage ? JSON.stringify({ serverRelativeUrl: evPhoto(['#1d4d9a', '#0f6b63', '#9b2c52', '#8a5a00'][id % 4], '#082244') }) : null
});
const evRows = [
  evRow(20, 'Summer BBQ (past)', '2026-09-12T10:00:00Z', '2026-09-12T14:00:00Z', false, 'Company', 'Roof terrace', true, ''),
  evRow(21, 'Quarterly Review (past)', '2026-10-05T08:00:00Z', '2026-10-05T09:00:00Z', false, 'Company', 'Online', false, ''),
  evRow(1, 'Blood Donation', '2026-10-30T07:00:00Z', '2026-10-30T11:00:00Z', false, 'Volunteering', 'B51 area (2nd floor, SPS kitchen)', true, '<p>Give blood and save lives. No appointment needed.</p>'),
  evRow(2, 'Safety Drill — Piraeus office', '2026-10-30T09:00:00Z', '2026-10-30T10:00:00Z', false, 'Safety', 'Piraeus office, lobby', true, '<p>Everyone meets at the <b>main lobby</b>.</p>'),
  evRow(3, 'Quarterly Town Hall', '2026-10-31T08:00:00Z', '2026-10-31T09:30:00Z', false, 'Company', 'Online (Teams)', false, '<p>Results and priorities for Q4.</p>'),
  evRow(4, 'Beach Clean-up', '2026-11-03T06:00:00Z', '2026-11-03T09:00:00Z', false, 'Volunteering', 'Glyfada beach', true, '<p>Gloves and bags are provided.</p>'),
  evRow(5, 'Athens Authentic Marathon', '2026-11-06T22:00:00Z', '2026-11-08T21:59:00Z', true, 'Sports', 'Athens', true, '<p>42km & 10km Sunday, 8 November</p><p>5km Saturday, 7 November</p>'),
  evRow(6, 'Wellbeing Workshop', '2026-11-12T10:00:00Z', '2026-11-12T12:00:00Z', false, 'Wellbeing', 'Meeting room 3', true, '<p>Stress and sleep, with a guest speaker.</p>'),
  evRow(7, 'Crew Family Day', '2026-11-14T07:00:00Z', '2026-11-14T15:00:00Z', false, 'Company', 'Kifissia', true, '<p>Games and lunch for the families of our seafarers.</p>'),
  evRow(8, 'Cyber Security Training', '2026-11-18T08:00:00Z', '2026-11-18T09:00:00Z', false, 'Training', 'Online (Teams)', false, '<p>Mandatory for all staff.</p>'),
  evRow(9, 'Charity Run', '2026-11-22T07:30:00Z', '2026-11-22T10:00:00Z', false, 'Sports', 'Pedion tou Areos', true, '<p>Register with HR.</p>'),
  evRow(10, 'Christmas Concert', '2026-12-04T17:00:00Z', '2026-12-04T19:00:00Z', false, 'Culture', 'Concert hall', true, '<p>Tickets at reception.</p>'),
  evRow(11, 'Year-end Dinner', '2026-12-18T18:00:00Z', '2026-12-18T22:00:00Z', false, 'Company', 'To be announced', true, '<p>Details soon.</p>'),
  evRow(12, 'Office Closed', '2026-12-25T00:00:00Z', '2026-12-25T21:59:00Z', true, 'Holiday', '', false, ''),
  evRow(13, 'Sustainability Webinar', '2026-12-09T09:00:00Z', '2026-12-09T10:00:00Z', false, 'Training', 'Online (Teams)', true, '<p>Our 2027 targets.</p>'),
  evRow(14, 'New Year Reception', '2027-01-08T09:00:00Z', '2027-01-08T11:00:00Z', false, 'Company', 'Head office', true, '<p>Welcome back drinks.</p>')
];
const fakeEventsListClient = {
  get: async (url: string) => {
    const u = decodeURIComponent(url);
    if (u.indexOf('/fields?') !== -1) return { ok: true, status: 200, json: async () => ({ value: ['Id', 'Title', 'EventDate', 'EndDate', 'Location', 'Description', 'fAllDayEvent', 'Category', 'BannerUrl'].map(n => ({ InternalName: n })) }) };
    const one = /items\((\d+)\)/.exec(u);
    if (one) { const row = evRows.find(r => r.Id === Number(one[1])); return row ? { ok: true, status: 200, json: async () => row } : { ok: false, status: 404 }; }
    if (u.indexOf('/items?') !== -1) {
      const top = /\$top=(\d+)/.exec(u); const rows = u.indexOf('$filter') !== -1 ? evRows.filter(r => new Date(r.EventDate).getTime() >= new Date(new Date().setHours(0, 0, 0, 0)).getTime()) : evRows;
      return { ok: true, status: 200, json: async () => ({ value: rows.slice(0, top ? Number(top[1]) : 100) }) };
    }
    return { ok: false, status: 404 };
  }
};
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
  ['Okyroe', 'Oil Tanker', 74999, '2004', 'Hyundai Heavy Industries', 'LR1'], ['Res Cogitans', 'Oil Tanker', 74999, '2004', 'Hyundai Heavy Industries', 'LR1'],
  ['Orwell', 'Oil / Chemical Tanker', 51745, '2010', 'Hyunday Mipo Dockyard - Ulsan South Korea', 'MR'], ['Jane', 'Oil / Chemical Tanker IMO III', 51505, '2008', 'STX Ship Building Co LTD', 'MR'],
  ['Janine K.', 'Oil Tanker', 48045, '2009', 'Japan', 'MR'], ['Joyce', 'Oil Tanker', 47344, '2007', 'Onomichi Dockyard CO LTD', 'MR'],
  ['Evgenia S.', 'Oil / Chemical Tanker IMO III', 46609, '2008', 'South Korea', 'MR'], ['Yume', 'Oil Tanker', 45994, '2010', 'Shin Kurushima Japan', 'MR'],
  ['Twinkle Star', 'Oil Tanker', 45750, '2006', 'Minaminippon Shipbuilding CO. LTD', 'MR']
].map((v, i) => ({ Id: i + 1, Title: v[0], VesselType: v[1], Dwt: v[2], BuiltYear: v[3], Shipyard: v[4], VesselClass: v[5] }));
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
      if (url.indexOf('VesselClass') !== -1 && new URLSearchParams(window.location.search).get('noClass') === '1') {
        return { ok: false, status: 400, json: async () => ({}) };
      }
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
ReactDom.render(React.createElement(FleetTable, { spHttpClient: fakeListClient as never, siteUrl: 'https://example.sharepoint.com/sites/demo', sortBy: 'dwt', sortDescending: true, maxRows: 0, showFilters: true, showSummary: true, summaryNoun: 'TANKERS', canManageLists: true }), newRoot());

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
    React.createElement('div', { className: 'zone' }, React.createElement(EventsWidget, { spHttpClient: (useFakeEvents ? fakeEventsClient : fakeSpHttpClient) as never, siteUrl: 'https://example.sharepoint.com/sites/demo' })),
    React.createElement(
      'div',
      { className: 'zone', style: { flex: 1, minWidth: 0, display: 'flex', flexDirection: 'column' } },
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
// events.html / event.html only: the five Events web parts
const evSite = 'https://example.sharepoint.com/sites/demo';
const evList = 'dd3316a0-3bc0-4d4c-acff-71851299dad7';
const evAt = (id: string): HTMLElement | null => document.getElementById(id);
const evRoot = (id: string, el: React.ReactElement): void => { const r = evAt(id); if (r) { ReactDom.render(el, r); } };
evRoot('ev-heading-root', React.createElement(SectionHeading, { title: 'Coming up', linkText: 'See all', linkUrl: '/sites/demo/SitePages/Events.aspx', showArrow: true, newTab: false, titleSize: 24, isEditMode: false }));
evRoot('ev-grid-root', React.createElement(EventsGrid, { spHttpClient: fakeEventsListClient as never, siteUrl: evSite, listRef: evList, eventPageUrl: '/event.html', maxEvents: 48, emptyText: 'No upcoming events.' }));
evRoot('ev-image-root', React.createElement(EventImage, { spHttpClient: fakeEventsListClient as never, siteUrl: evSite, listRef: evList, aspectRatio: '5 / 3', radius: 0, isEditMode: false }));
evRoot('ev-details-root', React.createElement(EventDetails, { spHttpClient: fakeEventsListClient as never, siteUrl: evSite, listRef: evList, showTitle: true, titleSize: 40, showWhen: true, showTimezone: true, showCategory: true, showLocation: true, showAbout: true, isEditMode: false }));
evRoot('ev-button-root', React.createElement(LinkButton, { label: 'See All Events', url: '/events.html', showArrow: true, newTab: false, color: '#A5192D', textColor: '#FFFFFF', align: 'left', spaceAbove: 40 }));

// box.html only: two sample Background Box web parts
const boxDefaults = { widthMode: 'column' as const, widthValue: 600, widthUnit: 'px' as const, align: 'left' as const, heightValue: 240, heightUnit: 'px' as const, color: '#082244', useGradient: false, colorEnd: '#123963', gradientAngle: 135, opacity: 100, radius: 30, perCorner: false, radiusTopLeft: 30, radiusTopRight: 30, radiusBottomRight: 30, radiusBottomLeft: 30, borderWidth: 0, borderColor: '#082244', shadow: false, pullUp: 0, isEditMode: false };
const boxA = document.getElementById('box-a-root');
if (boxA) { ReactDom.render(React.createElement(BackgroundBox, { ...boxDefaults, widthMode: 'full', useGradient: true, radius: 0, heightValue: 220, pullUp: 120 }), boxA); }
const boxB = document.getElementById('box-b-root');
if (boxB) { ReactDom.render(React.createElement(BackgroundBox, { ...boxDefaults, widthMode: 'custom', widthValue: 60, widthUnit: '%', align: 'center', color: '#C7D7E6', perCorner: true, radiusTopLeft: 80, radiusTopRight: 10, radiusBottomRight: 80, radiusBottomLeft: 10, shadow: true, borderWidth: 3, borderColor: '#082244', heightValue: 160 }), boxB); }
const mpSite = 'https://example.sharepoint.com/sites/demo';
ReactDom.render(React.createElement(MarketplaceFilters, { spHttpClient: fakeMarketplaceClient as never, siteUrl: mpSite, canManageLists: true }), document.getElementById('mp-filters-root'));
ReactDom.render(React.createElement(MarketplacePost, { spHttpClient: fakeMarketplaceClient as never, siteUrl: mpSite, canManageLists: true, eyebrow: 'SELL · GIVE · SWAP', heading: "Got something you don't need anymore? Post it in under a minute.", buttonLabel: 'Post an item', maxPhotos: 5 }), document.getElementById('mp-post-root'));
ReactDom.render(React.createElement(MarketplaceInfo, { spHttpClient: fakeMarketplaceClient as never, siteUrl: mpSite, isEditMode: false, rulesTitle: 'MARKETPLACE RULES', rulesText: 'Personal items only, no running a business\nNothing prohibited or unsafe\nListings expire after 30 days\nDeals are between colleagues; the company is not responsible for transactions', contactLabel: 'Questions? Contact HR', contactEmail: 'hr@example.com', contactSubject: 'Marketplace question', contactBody: 'Hello HR,\n\nI have a question about the Marketplace:\n\n', manageUrl: '', expiryDays: 30 }), document.getElementById('mp-info-root'));
ReactDom.render(React.createElement(MarketplaceListings, { spHttpClient: fakeMarketplaceClient as never, siteUrl: mpSite, canManageLists: true, expiryDays: 30 }), document.getElementById('mp-listings-root'));
ReactDom.render(
  React.createElement(NewsGrid, {
    spHttpClient: fakeHomeClient as never,
    siteUrl: 'https://example.sharepoint.com/sites/demo',
    // ?gridPinned=N (0-8) pins that many sample pages in the News Grid; default 1.
    pinnedNews: Array.from({ length: Number(new URLSearchParams(window.location.search).get('gridPinned') ?? 1) }, (_, i) => 'Pinned-' + (i + 1) + '.aspx')
  }),
  document.getElementById('news-grid-root')
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
