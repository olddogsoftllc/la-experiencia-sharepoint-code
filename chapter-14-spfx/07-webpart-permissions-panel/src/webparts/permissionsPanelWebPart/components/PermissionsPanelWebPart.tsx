import * as React from 'react';
import {
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Stack,
  PrimaryButton
} from '@fluentui/react';
import styles from './PermissionsPanelWebPart.module.scss';
import type { IPermissionsPanelWebPartProps } from './IPermissionsPanelWebPartProps';
import { useGraphQuery } from '../hooks/useGraphQuery';
import { ISitePermission } from '../models/ISitePermission';

const PermissionsPanel: React.FunctionComponent<IPermissionsPanelWebPartProps> = (props) => {
  const {
    graphService,
    siteId,
    isDarkTheme,
    hasTeamsContext,
    userDisplayName
  } = props;

  const { data: perms, loading, error, refetch } = useGraphQuery<ISitePermission[]>(
    () => graphService.getSitePermissions(siteId),
    [siteId]
  );

  return (
    <section className={`${styles.permissionsPanelWebPart} ${hasTeamsContext ? styles.teams : ''} ${isDarkTheme ? styles.dark : ''}`}>
      <div className={styles.welcome}>
        <h2>Permissions Panel, {userDisplayName}</h2>
        <div className={styles.subtitle}>Reading <strong>/sites/{siteId}/permissions</strong></div>
      </div>

      <Stack tokens={{ childrenGap: 12 }}>
        <PrimaryButton text="Refresh" onClick={() => refetch()} />

        {loading && <Spinner label="Querying site permissions..." size={SpinnerSize.large} />}

        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline>
            Cannot read site permissions: {error.message}.
            <br />
            <strong>/sites/&#123;id&#125;/permissions</strong> is an admin operation: it typically
            needs admin consent and returns <strong>403</strong> under Sites.Selected
            (see chapter 5 and the book&rsquo;s &ldquo;least privilege&rdquo; rule).
          </MessageBar>
        )}

        {!loading && !error && perms && perms.length === 0 && (
          <MessageBar>No permission entries returned.</MessageBar>
        )}

        {!loading && !error && perms && perms.length > 0 && (
          <Stack tokens={{ childrenGap: 8 }}>
            {perms.map((p: ISitePermission) => (
              <div key={p.id} className={styles.permRow}>
                <div className={styles.roles}>
                  {p.roles.map((r: string) => (
                    <span key={r} className={styles.role}>{r}</span>
                  ))}
                </div>
                <div>
                  Granted to: <strong>{p.grantedTo}</strong>
                  {p.isInherited ? ' (inherited)' : ''}
                </div>
              </div>
            ))}
          </Stack>
        )}
      </Stack>
    </section>
  );
};

export default PermissionsPanel;