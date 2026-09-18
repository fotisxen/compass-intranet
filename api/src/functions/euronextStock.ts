import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';

// Star Bulk's own site (starbulk.com) has no Access-Control-Allow-Origin
// header on this endpoint, so the browser blocks it as a cross-origin
// request when called directly from our SharePoint site — this proxies it
// server-side (no CORS restriction between servers) the same way fleet.ts
// proxies the fleet positions API. No API key needed: this is the same
// public endpoint starbulk.com's own homepage calls client-side.
const EURONEXT_STOCK_FEED_URL = 'https://www.starbulk.com/?module=investorrelations&action=get-euronext-stock-feed';

function corsHeaders(): Record<string, string> {
  const allowedOrigin = process.env.ALLOWED_ORIGIN || '*';
  return {
    'Access-Control-Allow-Origin': allowedOrigin,
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };
}

export async function euronextStock(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  if (request.method === 'OPTIONS') {
    return { status: 204, headers: corsHeaders() };
  }

  let response: Response;
  try {
    response = await fetch(EURONEXT_STOCK_FEED_URL);
  } catch (err) {
    context.error('Failed to reach starbulk.com for the Euronext stock feed', err);
    return { status: 502, headers: corsHeaders(), jsonBody: { error: 'Could not reach the Euronext stock feed.' } };
  }

  if (!response.ok) {
    context.error('Euronext stock feed request failed', response.status);
    return { status: 502, headers: corsHeaders(), jsonBody: { error: 'The Euronext stock feed returned an error.' } };
  }

  const data = await response.json();
  return { headers: corsHeaders(), jsonBody: data };
}

app.http('euronextStock', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: euronextStock
});
