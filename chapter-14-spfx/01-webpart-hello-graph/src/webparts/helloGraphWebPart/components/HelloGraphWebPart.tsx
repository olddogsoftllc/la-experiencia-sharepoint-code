import * as React from 'react';
import {
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Stack,
  DocumentCard,
  DocumentCardTitle,
  DocumentCardLocation,
  PrimaryButton
} from '@fluentui/react';
import styles from './HelloGraphWebPart.module.scss';
import type { IHelloGraphWebPartProps } from './IHelloGraphWebPartProps';
import { useGraphQuery } from '../hooks/useGraphQuery';
import { ISite } from '../models/ISite';

const HelloGraph: React.FunctionComponent<IHelloGraphWebPartProps> = (props) => {
  const {
    graphService,
    query,
    maxResults,
    isDarkTheme,
    hasTeamsContext,
    userDisplayName
  } = props;

  const { data: sites, loading, error, refetch } = useGraphQuery<ISite[]>(
    () => graphService.searchSites(query, maxResults),
    [query, maxResults]
  );

  return (
    <section className={`${styles.helloGraphWebPart} ${hasTeamsContext ? styles.teams : ''} ${isDarkTheme ? styles.dark : ''}`}>
      <div className={styles.welcome}>
        <h2>Hello Graph, {userDisplayName}!</h2>
        <div className={styles.subtitle}>
          Querying <strong>/sites?search={query || ''}</strong> (top {maxResults})
        </div>
      </div>

      <Stack tokens={{ childrenGap: 12 }}>
        <PrimaryButton text="Refresh" onClick={() => refetch()} />

        {loading && <Spinner label="Querying Microsoft Graph..." size={SpinnerSize.large} />}

        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
            Graph error: {error.message}. Verify the <strong>Sites.Read.All</strong> permission is
            approved in SharePoint Admin Center → API Access.
          </MessageBar>
        )}

        {!loading && !error && sites && sites.length === 0 && (
          <MessageBar>No sites matched &lsquo;{query}&rsquo;.</MessageBar>
        )}

        {!loading && !error && sites && sites.length > 0 && (
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
            {sites.map((s: ISite) => (
              <DocumentCard key={s.id} onClickHref={s.url}>
                <DocumentCardTitle title={s.displayName || s.name || s.url} shouldTruncate />
                <DocumentCardLocation location={s.url} />
              </DocumentCard>
            ))}
          </Stack>
        )}
      </Stack>
    </section>
  );
};

export default HelloGraph;