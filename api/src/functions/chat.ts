import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';

interface IChatMessage {
  role: 'user' | 'assistant';
  content: string;
}

interface IChatRequestBody {
  messages: IChatMessage[];
}

const MAX_MESSAGES = 20;
const MAX_MESSAGE_LENGTH = 4000;

function corsHeaders(): Record<string, string> {
  const allowedOrigin = process.env.ALLOWED_ORIGIN || '*';
  return {
    'Access-Control-Allow-Origin': allowedOrigin,
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };
}

function isValidMessages(messages: unknown): messages is IChatMessage[] {
  if (!Array.isArray(messages) || messages.length === 0 || messages.length > MAX_MESSAGES) {
    return false;
  }
  return messages.every(
    m =>
      m &&
      (m.role === 'user' || m.role === 'assistant') &&
      typeof m.content === 'string' &&
      m.content.length > 0 &&
      m.content.length <= MAX_MESSAGE_LENGTH
  );
}

export async function chat(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  if (request.method === 'OPTIONS') {
    return { status: 204, headers: corsHeaders() };
  }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    context.error('OPENAI_API_KEY is not configured on the function app.');
    return { status: 500, headers: corsHeaders(), jsonBody: { error: 'Chat is not configured on the server.' } };
  }

  let body: IChatRequestBody;
  try {
    body = (await request.json()) as IChatRequestBody;
  } catch {
    return { status: 400, headers: corsHeaders(), jsonBody: { error: 'Invalid JSON body.' } };
  }

  if (!isValidMessages(body.messages)) {
    return {
      status: 400,
      headers: corsHeaders(),
      jsonBody: { error: `messages must be a non-empty array of at most ${MAX_MESSAGES} user/assistant messages.` }
    };
  }

  const systemMessage = {
    role: 'system' as const,
    content: 'You are a helpful assistant embedded in a SharePoint intranet page. Keep answers concise.'
  };

  let response: Response;
  try {
    response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: process.env.OPENAI_MODEL || 'gpt-4o-mini',
        messages: [systemMessage, ...body.messages]
      })
    });
  } catch (err) {
    context.error('Failed to reach OpenAI', err);
    return { status: 502, headers: corsHeaders(), jsonBody: { error: 'Could not reach the AI provider.' } };
  }

  if (!response.ok) {
    const errorText = await response.text();
    context.error('OpenAI request failed', response.status, errorText);
    return { status: 502, headers: corsHeaders(), jsonBody: { error: 'The AI provider returned an error.' } };
  }

  const data = (await response.json()) as { choices: { message: { content: string } }[] };
  const reply = data.choices?.[0]?.message?.content ?? '';

  return { headers: corsHeaders(), jsonBody: { reply } };
}

app.http('chat', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: chat
});
