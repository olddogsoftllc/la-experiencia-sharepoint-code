import { AadHttpClient } from '@microsoft/sp-http';

/**
 * Calls a custom Azure AD-protected API via AadHttpClient.
 *
 * SPFx obtains the token automatically through `aadHttpClientFactory` — the
 * API's App ID URI is declared as a webApiPermissionRequest in
 * package-solution.json and approved in SharePoint Admin Center → API Access.
 *
 * Covers the book's cap14 "MSGraphClient vs AadHttpClient".
 */
export class OwnApiService {
  constructor(
    private readonly client: AadHttpClient,
    private readonly apiUrl: string
  ) {}

  public async getData(): Promise<unknown> {
    const response = await this.client.get(this.apiUrl, AadHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(`API returned ${response.status} ${response.statusText}`);
    }
    return response.json();
  }
}