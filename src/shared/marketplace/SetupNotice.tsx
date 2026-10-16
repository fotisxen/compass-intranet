import * as React from 'react';
import styles from './Marketplace.module.scss';

export interface ISetupNoticeProps {
  /** Whether the signed-in user can create lists — only then is the setup button offered. */
  canManageLists: boolean;
  /** Creates the Marketplace lists. */
  onSetup: () => Promise<void>;
}

/** Shown instead of a Marketplace web part until the lists exist: owners get a button, everyone else a short note. */
export default function SetupNotice(props: ISetupNoticeProps): React.ReactElement {
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>();

  const run = (): void => {
    setBusy(true);
    setError(undefined);
    props.onSetup().catch((e: Error) => {
      setError(e.message);
      setBusy(false);
    });
  };

  return (
    <div className={styles.notice}>
      {props.canManageLists ? (
        <>
          The Marketplace lists (&quot;Compass Marketplace&quot; and &quot;Compass Marketplace Categories&quot;) do not exist on this site yet.
          <div className={styles.setupButton}>
            <button type="button" className={styles.pill} onClick={run} disabled={busy}>
              {busy ? 'Creating…' : 'Create the Marketplace lists'}
            </button>
          </div>
          {error && <div className={styles.noticeError}>{error}</div>}
        </>
      ) : (
        <>The Marketplace is not set up yet. Please ask a site owner to open this page once.</>
      )}
    </div>
  );
}
