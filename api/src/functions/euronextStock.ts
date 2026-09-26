import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';

// The Euronext Athens quote comes from the inbroker market-data feed. That
// feed sends no Access-Control-Allow-Origin header, so the browser blocks it
// when called directly from our SharePoint site — this proxies it
// server-side (no CORS restriction between servers) the same way fleet.ts
// proxies the fleet positions API. The full URL (it embeds a session id and
// user name) lives in the EURONEXT_FEED_URL app setting, not in source
// control, e.g.:
//   https://fc1a.inbroker.com/Info?userName=...&IBSessionId=...&company=...&lang=EN&format=json&code=SBLK.ATH

function corsHeaders(): Record<string, string> {
  const allowedOrigin = process.env.ALLOWED_ORIGIN || '*';
  return {
    'Access-Control-Allow-Origin': allowedOrigin,
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };
}

interface IInbrokerRow {
  price?: number;
  pricePrevClosePricePDelta?: number;
  currCode?: string;
}

interface IInbrokerResponse {
  'inbroker-transactions'?: { row?: IInbrokerRow };
}

export async function euronextStock(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  if (request.method === 'OPTIONS') {
    return { status: 204, headers: corsHeaders() };
  }

  const feedUrl = process.env.EURONEXT_FEED_URL;
  if (!feedUrl) {
    context.error('EURONEXT_FEED_URL is not configured on the function app.');
    return { status: 500, headers: corsHeaders(), jsonBody: { error: 'The Euronext feed is not configured on the server.' } };
  }

  let response: Response;
  try {
    response = await fetch(feedUrl);
  } catch (err) {
    context.error('Failed to reach the Euronext feed', err);
    return { status: 502, headers: corsHeaders(), jsonBody: { error: 'Could not reach the Euronext feed.' } };
  }

  if (!response.ok) {
    context.error('Euronext feed request failed', response.status);
    return { status: 502, headers: corsHeaders(), jsonBody: { error: 'The Euronext feed returned an error.' } };
  }

  const data = (await response.json()) as IInbrokerResponse;
  const row = data['inbroker-transactions']?.row;
  if (!row || typeof row.price !== 'number' || typeof row.pricePrevClosePricePDelta !== 'number') {
    // An expired session id comes back as a well-formed reply with no quote.
    context.error('Euronext feed reply had no quote — the session id in EURONEXT_FEED_URL may have expired.');
    return { status: 502, headers: corsHeaders(), jsonBody: { error: 'The Euronext feed returned no quote.' } };
  }

  return {
    headers: corsHeaders(),
    jsonBody: {
      symbol: 'SBLK',
      exchange: 'EURONEXT',
      currency: row.currCode || 'EUR',
      lastTrade: String(row.price),
      changePercent: String(row.pricePrevClosePricePDelta)
    }
  };
}

app.http('euronextStock', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: euronextStock
});
