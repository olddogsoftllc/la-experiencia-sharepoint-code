import type { GraphService } from '../services/GraphService';

export interface IHelloGraphWebPartProps {
  graphService: GraphService;
  query: string;
  maxResults: number;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
}