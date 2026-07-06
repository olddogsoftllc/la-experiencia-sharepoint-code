import * as React from 'react';
import {
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Stack,
  PrimaryButton
} from '@fluentui/react';
import styles from './OwnApiWebPart.module.scss';
import type { IOwnApiWebPartProps } from './IOwnApiWebPartProps';
import { useAsync } from '../hooks/useAsync';

const OwnApiWebPart: React.FunctionComponent<IOwnApiWebPartProps> = (props) => {
  const {
    apiService,
    isDarkTheme,
    hasTeamsContext,
    userDisplayName
  } = props;

  const { data, loading, error, refetch } = useAsync<unknown>(
    () => apiService ? apiService.getData() : Promise.reject(new Error('API service not initialized')),
    [apiService]
  );

  return (
    <section className={`${styles.ownApiWebPart} ${hasTeamsContext ? styles.teams : ''} ${isDarkTheme ? styles.dark : ''}`}>
      <div className={styles.welcome}>
        <h2>Own API Client, {userDisplayName}</h2>
        <div className={styles.subtitle}>Calling a custom Azure AD-protected API via <strong>AadHttpClient</strong></div>
      </div>

      <Stack tokens={{ childrenGap: 12 }}>
        <PrimaryButton text="Refresh" onClick={() => refetch()} />

        {loading && <Spinner label="Calling API..." size={SpinnerSize.large} />}

        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
            API error: {error.message}. Approve the API&rsquo;s App ID URI in SharePoint Admin Center → API Access.
          </MessageBar>
        )}

        {!loading && !error && data !== undefined && (
          <pre className={styles.json}>{JSON.stringify(data, null, 2)}</pre>
        )}
      </Stack>
    </section>
  );
};

export default OwnApiWebPart;