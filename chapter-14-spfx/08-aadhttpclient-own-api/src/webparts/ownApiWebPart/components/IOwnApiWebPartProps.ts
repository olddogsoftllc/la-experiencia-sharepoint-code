import type { OwnApiService } from '../services/OwnApiService';

export interface IOwnApiWebPartProps {
  apiService: OwnApiService | undefined;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
}