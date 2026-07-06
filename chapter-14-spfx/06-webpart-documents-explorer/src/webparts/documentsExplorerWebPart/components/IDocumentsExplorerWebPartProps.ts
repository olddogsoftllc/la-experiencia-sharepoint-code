import type { GraphService } from '../services/GraphService';

export interface IDocumentsExplorerWebPartProps {
  graphService: GraphService;
  siteId: string;
  maxResults: number;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
}