import * as React from 'react';
import { Stack, Text } from '@fluentui/react';
import styles from './TeamsEnabledWebPart.module.scss';
import type { ITeamsEnabledWebPartProps } from './ITeamsEnabledWebPartProps';

const TeamsEnabledWebPart: React.FunctionComponent<ITeamsEnabledWebPartProps> = (props) => {
  const {
    title,
    host,
    theme,
    isDarkTheme,
    hasTeamsContext
  } = props;

  return (
    <section className={`${styles.teamsEnabledWebPart} ${hasTeamsContext ? styles.teams : ''} ${isDarkTheme ? styles.dark : ''}`}>
      <div className={styles.welcome}>
        <h2>{title}</h2>
      </div>

      <Stack tokens={{ childrenGap: 8 }}>
        <Text>
          Host: <strong>{host}</strong>
          {hasTeamsContext ? ' (running inside Microsoft Teams)' : ' (running in SharePoint)'}
        </Text>
        <Text>Theme: <strong>{theme}</strong>{isDarkTheme ? ' (inverted)' : ''}</Text>
        <Text className={styles.note}>
          The same SPFx bundle renders here and as a Teams tab — see <code>teams/</code> icons and
          the manifest&rsquo;s <code>supportedHosts</code>.
        </Text>
      </Stack>
    </section>
  );
};

export default TeamsEnabledWebPart;