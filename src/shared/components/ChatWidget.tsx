import * as React from 'react';
import { AadHttpClient, type AadHttpClientFactory } from '@microsoft/sp-http';
import styles from './ChatWidget.module.scss';

export interface IChatCitation {
  title: string;
  url: string;
}

export interface IChatMessage {
  role: 'user' | 'assistant';
  content: string;
  citations?: IChatCitation[];
}

export interface IChatWidgetProps {
  /** Legacy/dev fallback: POST {messages} -> {reply} (e.g. our own OpenAI proxy in api/). Ignored once the AAD props below are set. */
  apiUrl?: string;
  /** From context.aadHttpClientFactory — lets the widget call the real AskStarBulk backend with the signed-in user's own token. */
  aadHttpClientFactory?: AadHttpClientFactory;
  /** AAD App ID URI of the AskStarBulk custom API, e.g. api://1bc53e14-c0fd-41d3-a318-ba88918961b1 */
  chatResourceUri?: string;
  /** Full query endpoint, e.g. https://func-askstarbulk-prod-.../api/query */
  chatQueryUrl?: string;
}

export interface IChatWidgetState {
  messages: IChatMessage[];
  input: string;
  isLoading: boolean;
  error?: string;
}

interface IAskStarBulkResponse {
  answer: string;
  citations?: IChatCitation[];
}

// Used only while no real assistant is configured — e.g. testing the chat
// panel's own layout and scroll behavior before a backend is reachable.
// Lengths are deliberately mixed (one-liners next to multi-line paragraphs)
// so a real conversation's worth of scrolling gets exercised.
const MOCK_REPLIES: string[] = [
  'Sure, I can help with that. What specifically are you trying to find?',
  "Good question! Let me walk you through it.\n\nStar Bulk's fleet operates across several vessel classes, and the Fleet section under Newsroom has the latest updates on individual ship movements and port calls.",
  "I don't have a definitive answer for that one, but the HR team on the My Workplace page should be able to help.",
  "Here's a quick summary:\n\n1. Check the Public Holidays widget on the homepage for office closures.\n2. Vacation requests go through SAP HRMS.\n3. For anything urgent, reach out to your line manager directly.",
  'Got it — noted!',
  "That depends on your department. Fleet Operations and Chartering each have slightly different processes, so it's worth checking the relevant policy under Company Profile → Policies.",
  'Absolutely, happy to help with that.',
  "Thanks for the context. I'd recommend starting with the Digital Hub page — it links out to most of the internal tools you'd need, including SAP and the IT Helpdesk.",
  "I'm a mock reply right now — the real assistant isn't connected in this environment yet, but the chat panel itself is fully wired up and ready.",
  'Let me know if there’s anything else I can help you find on the intranet!'
];

// Renders the small subset of markdown the real assistant actually sends
// (#/##/### headings, **bold**, "- "/"* " bullets, "1. " numbered lists)
// without pulling in a markdown library or touching innerHTML.
function renderInline(line: string, keyPrefix: string): React.ReactNode[] {
  const parts = line.split(/(\*\*[^*]+\*\*)/g);
  return parts.map((part, i) => {
    if (part.startsWith('**') && part.endsWith('**') && part.length > 3) {
      return <strong key={`${keyPrefix}-${i}`}>{part.slice(2, -2)}</strong>;
    }
    return <React.Fragment key={`${keyPrefix}-${i}`}>{part}</React.Fragment>;
  });
}

function renderMarkdownLite(text: string): React.ReactNode {
  const lines = text.split('\n');
  return lines.map((line, i) => {
    const key = `md-${i}`;
    const heading = /^#{1,6}\s+(.*)$/.exec(line);
    if (heading) {
      return (
        <div key={key} className={styles.mdHeading}>
          {renderInline(heading[1], key)}
        </div>
      );
    }
    const bullet = /^[-*]\s+(.*)$/.exec(line);
    if (bullet) {
      return (
        <div key={key} className={styles.mdListItem}>
          • {renderInline(bullet[1], key)}
        </div>
      );
    }
    const numbered = /^(\d+)\.\s+(.*)$/.exec(line);
    if (numbered) {
      return (
        <div key={key} className={styles.mdListItem}>
          {numbered[1]}. {renderInline(numbered[2], key)}
        </div>
      );
    }
    if (line.trim() === '') {
      return <div key={key} className={styles.mdSpacer} />;
    }
    return <div key={key}>{renderInline(line, key)}</div>;
  });
}

export default class ChatWidget extends React.Component<IChatWidgetProps, IChatWidgetState> {
  private _mockReplyIndex = 0;

  constructor(props: IChatWidgetProps) {
    super(props);
    this.state = { messages: [], input: '', isLoading: false };
  }

  private _onInputChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ input: e.target.value });
  };

  private _onKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter') {
      this._send();
    }
  };

  private _hasAadBackend(): boolean {
    const { aadHttpClientFactory, chatResourceUri, chatQueryUrl } = this.props;
    return !!aadHttpClientFactory && !!chatResourceUri && !!chatQueryUrl;
  }

  private _send = (): void => {
    const trimmed = this.state.input.trim();
    if (!trimmed || this.state.isLoading) {
      return;
    }

    const nextMessages: IChatMessage[] = [...this.state.messages, { role: 'user', content: trimmed }];
    this.setState({ messages: nextMessages, input: '', isLoading: true, error: undefined });

    if (this._hasAadBackend()) {
      this._sendAad(trimmed);
      return;
    }

    if (!this.props.apiUrl) {
      this._sendMock();
      return;
    }

    fetch(this.props.apiUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ messages: nextMessages })
    })
      .then(response => {
        if (!response.ok) {
          throw new Error(`Request failed with status ${response.status}`);
        }
        return response.json();
      })
      .then((data: { reply: string }) => {
        this.setState(prev => ({
          messages: [...prev.messages, { role: 'assistant', content: data.reply }],
          isLoading: false
        }));
      })
      .catch(() => {
        this.setState({
          isLoading: false,
          error: 'Something went wrong reaching the assistant. Please try again.'
        });
      });
  };

  // Same account/UI, real backend: acquire a token for the AskStarBulk API
  // via SPFx's own AAD broker (no key in our code) and ask it the question.
  private _sendAad(query: string): void {
    const { aadHttpClientFactory, chatResourceUri, chatQueryUrl } = this.props;

    aadHttpClientFactory!
      .getClient(chatResourceUri!)
      .then(client =>
        client.post(chatQueryUrl!, AadHttpClient.configurations.v1, {
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ query })
        })
      )
      .then(response => {
        if (!response.ok) {
          throw new Error(`Request failed with status ${response.status}`);
        }
        return response.json();
      })
      .then((data: IAskStarBulkResponse) => {
        this.setState(prev => ({
          messages: [...prev.messages, { role: 'assistant', content: data.answer, citations: data.citations }],
          isLoading: false
        }));
      })
      .catch(() => {
        this.setState({
          isLoading: false,
          error: 'Something went wrong reaching the assistant. Please try again.'
        });
      });
  }

  // No real assistant configured — reply from the local pool instead of
  // erroring, so the panel can be exercised end to end before the backend
  // is deployed.
  private _sendMock(): void {
    const reply = MOCK_REPLIES[this._mockReplyIndex % MOCK_REPLIES.length];
    this._mockReplyIndex++;

    const delay = 500 + Math.random() * 700;
    setTimeout(() => {
      this.setState(prev => ({
        messages: [...prev.messages, { role: 'assistant', content: reply }],
        isLoading: false
      }));
    }, delay);
  }

  public render(): React.ReactElement<IChatWidgetProps> {
    const { messages, input, isLoading, error } = this.state;

    return (
      <div className={styles.chatWidget}>
        <h3>Ask the assistant</h3>
        <div className={styles.messages}>
          {messages.length === 0 && <div className={styles.placeholder}>Ask a question to get started.</div>}
          {messages.map((m, i) => (
            <div key={i} className={m.role === 'user' ? styles.userMessage : styles.assistantMessage}>
              {m.role === 'assistant' ? renderMarkdownLite(m.content) : m.content}
              {m.citations && m.citations.length > 0 && (
                <div className={styles.citations}>
                  {m.citations.map((c, ci) => (
                    <a key={ci} href={c.url} target="_blank" rel="noopener noreferrer" className={styles.citationLink}>
                      {c.title}
                    </a>
                  ))}
                </div>
              )}
            </div>
          ))}
          {isLoading && <div className={styles.assistantMessage}>Thinking…</div>}
        </div>
        {error && <div className={styles.error}>{error}</div>}
        <div className={styles.inputRow}>
          <input
            type="text"
            className={styles.input}
            value={input}
            placeholder="Type a message..."
            onChange={this._onInputChange}
            onKeyDown={this._onKeyDown}
            disabled={isLoading}
          />
          <button className={styles.sendButton} onClick={() => this._send()} disabled={isLoading}>
            Send
          </button>
        </div>
      </div>
    );
  }
}
