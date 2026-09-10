import * as React from 'react';
import styles from './ChatWidget.module.scss';

export interface IChatMessage {
  role: 'user' | 'assistant';
  content: string;
}

export interface IChatWidgetProps {
  apiUrl: string;
}

export interface IChatWidgetState {
  messages: IChatMessage[];
  input: string;
  isLoading: boolean;
  error?: string;
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

  private _send = (): void => {
    const trimmed = this.state.input.trim();
    if (!trimmed || this.state.isLoading) {
      return;
    }

    if (!this.props.apiUrl) {
      this.setState({ error: 'Chat is not configured yet. Set the assistant API URL in the web part property pane.' });
      return;
    }

    const nextMessages: IChatMessage[] = [...this.state.messages, { role: 'user', content: trimmed }];
    this.setState({ messages: nextMessages, input: '', isLoading: true, error: undefined });

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

  public render(): React.ReactElement<IChatWidgetProps> {
    const { messages, input, isLoading, error } = this.state;

    return (
      <div className={styles.chatWidget}>
        <h3>Ask the assistant</h3>
        <div className={styles.messages}>
          {messages.length === 0 && <div className={styles.placeholder}>Ask a question to get started.</div>}
          {messages.map((m, i) => (
            <div key={i} className={m.role === 'user' ? styles.userMessage : styles.assistantMessage}>
              {m.content}
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
