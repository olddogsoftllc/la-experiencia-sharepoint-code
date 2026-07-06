// src/graphAuth.ts — reusable Graph factory (same pattern as chapter-02-auth common/)
// App-only with client secret. For production use a certificate (see chapter-02-auth).
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch'; // fetch polyfill for Node (provided by @azure/identity transitively)

let _client: Client | undefined;

/**
 * Lazily builds a Graph client in app-only mode.
 * Reuses the same pattern as the chapter 2 common module: env vars + credential.
 */
export function getGraphClient(): Client {
  if (_client) return _client;

  const tenantId = process.env.GRAPH_TENANT_ID;
  const clientId = process.env.GRAPH_CLIENT_ID;
  const clientSecret = process.env.GRAPH_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error(
      'GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET are missing in the environment. Copy .env.example to .env.'
    );
  }

  // TokenCredentialProvider accepted by @microsoft/microsoft-graph-client
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

  _client = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        const token = await credential.getToken('https://graph.microsoft.com/.default');
        return token.token;
      },
    },
  });

  return _client;
}

// Minimal types to type the responses (avoids `any` — consistent with chapter 14 GraphService)
export interface IGraphSite {
  id: string;
  displayName: string;
  webUrl: string;
}
export interface IGraphSiteCollection {
  value: IGraphSite[];
}
export interface IGraphDrive {
  id: string;
  name: string;
}
export interface IGraphDriveItem {
  id: string;
  name: string;
}