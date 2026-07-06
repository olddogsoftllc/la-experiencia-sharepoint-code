export interface ITeamsEnabledWebPartProps {
  title: string;
  host: string;          // SharePoint | Teams | Office | Outlook | Unknown
  theme: string;         // default | dark | contrast | (SharePoint theme)
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
}