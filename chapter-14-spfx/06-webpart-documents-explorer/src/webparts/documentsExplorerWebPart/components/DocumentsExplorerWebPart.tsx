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
import styles from './DocumentsExplorerWebPart.module.scss';
import type { IDocumentsExplorerWebPartProps } from './IDocumentsExplorerWebPartProps';
import { useGraphQuery } from '../hooks/useGraphQuery';
import { IDrive, IDriveItem } from '../models/IDrive';

const DocumentsExplorer: React.FunctionComponent<IDocumentsExplorerWebPartProps> = (props) => {
  const {
    graphService,
    siteId,
    maxResults,
    isDarkTheme,
    hasTeamsContext,
    userDisplayName
  } = props;

  // 1) libraries (drives) of the current site
  const drivesState = useGraphQuery<IDrive[]>(
    () => graphService.getSiteDrives(siteId),
    [siteId]
  );
  const firstDriveId: string | undefined = drivesState.data?.[0]?.id;

  // 2) root items of the first library (re-runs when firstDriveId is known)
  const itemsState = useGraphQuery<IDriveItem[]>(
    () => firstDriveId
      ? graphService.getDriveRootItems(firstDriveId, maxResults)
      : Promise.resolve([] as IDriveItem[]),
    [firstDriveId, maxResults]
  );

  return (
    <section className={`${styles.documentsExplorerWebPart} ${hasTeamsContext ? styles.teams : ''} ${isDarkTheme ? styles.dark : ''}`}>
      <div className={styles.welcome}>
        <h2>Documents Explorer, {userDisplayName}</h2>
      </div>

      <Stack tokens={{ childrenGap: 16 }}>
        <div>
          <h3 className={styles.sectionTitle}>Libraries in this site</h3>
          <PrimaryButton text="Refresh" onClick={() => { drivesState.refetch(); itemsState.refetch(); }} />

          {drivesState.loading && <Spinner label="Loading libraries..." size={SpinnerSize.large} />}
          {drivesState.error && (
            <MessageBar messageBarType={MessageBarType.error}>
              Graph error: {drivesState.error.message}. Approve <strong>Sites.Read.All</strong> /
              <strong>Files.Read.All</strong> in SharePoint Admin Center → API Access.
            </MessageBar>
          )}
          {!drivesState.loading && !drivesState.error && drivesState.data && drivesState.data.length === 0 && (
            <MessageBar>No libraries found.</MessageBar>
          )}
          {!drivesState.loading && !drivesState.error && drivesState.data && drivesState.data.length > 0 && (
            <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
              {drivesState.data.map((d: IDrive) => (
                <DocumentCard key={d.id} onClickHref={d.webUrl}>
                  <DocumentCardTitle title={d.name} shouldTruncate />
                  <DocumentCardLocation location={d.driveType ?? 'documentLibrary'} />
                </DocumentCard>
              ))}
            </Stack>
          )}
        </div>

        <div>
          <h3 className={styles.sectionTitle}>Root items of the first library</h3>
          {itemsState.loading && <Spinner label="Loading items..." size={SpinnerSize.medium} />}
          {itemsState.error && (
            <MessageBar messageBarType={MessageBarType.warning}>
              Could not load items: {itemsState.error.message}
            </MessageBar>
          )}
          {!itemsState.loading && !itemsState.error && itemsState.data && itemsState.data.length === 0 && (
            <MessageBar>No items in root.</MessageBar>
          )}
          {!itemsState.loading && !itemsState.error && itemsState.data && itemsState.data.length > 0 && (
            <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
              {itemsState.data.map((i: IDriveItem) => (
                <DocumentCard key={i.id} onClickHref={i.webUrl}>
                  <DocumentCardTitle title={`${i.isFolder ? '📁 ' : '📄 '}${i.name}`} shouldTruncate />
                </DocumentCard>
              ))}
            </Stack>
          )}
        </div>
      </Stack>
    </section>
  );
};

export default DocumentsExplorer;