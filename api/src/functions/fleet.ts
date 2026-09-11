import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';

function corsHeaders(): Record<string, string> {
  const allowedOrigin = process.env.ALLOWED_ORIGIN || '*';
  return {
    'Access-Control-Allow-Origin': allowedOrigin,
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };
}

export async function fleet(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  if (request.method === 'OPTIONS') {
    return { status: 204, headers: corsHeaders() };
  }

  const apiUrl = process.env.FLEET_API_URL;
  const apiKey = process.env.FLEET_API_KEY;
  if (!apiUrl || !apiKey) {
    context.error('FLEET_API_URL / FLEET_API_KEY are not configured on the function app.');
    return {
      status: 500,
      headers: corsHeaders(),
      jsonBody: { error: 'Fleet positions are not configured on the server.' }
    };
  }

  let response: Response;
  try {
    const separator = apiUrl.includes('?') ? '&' : '?';
    response = await fetch(`${apiUrl}${separator}code=${encodeURIComponent(apiKey)}`);
  } catch (err) {
    context.error('Failed to reach the fleet positions provider', err);
    return { status: 502, headers: corsHeaders(), jsonBody: { error: 'Could not reach the fleet positions provider.' } };
  }

  if (!response.ok) {
    context.error('Fleet positions request failed', response.status);
    return { status: 502, headers: corsHeaders(), jsonBody: { error: 'The fleet positions provider returned an error.' } };
  }

  const data = await response.json();
  return { headers: corsHeaders(), jsonBody: data };
}

app.http('fleet', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: fleet
});
