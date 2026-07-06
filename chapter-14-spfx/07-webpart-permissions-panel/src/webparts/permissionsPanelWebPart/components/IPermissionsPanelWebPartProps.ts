import type { GraphService } from '../services/GraphService';

export interface IPermissionsPanelWebPartProps {
  graphService: GraphService;
  siteId: string;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
}