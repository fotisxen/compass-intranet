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
      this.setState({
        isLoading: false,
        error: 'The assistant is not configured for this site yet.'
      });
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
