// src/graphAuth.ts — factory de Graph reutilizable (mismo patrón que chapter-02-auth common/)
// App-only con client secret. Para producción usa certificado (ver chapter-02-auth).
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch'; // polyfill fetch en Node (lo trae @azure/identity vía transitive)

let _client: Client | undefined;

/**
 * Construye (lazy) un cliente de Graph en modo app-only.
 * Reusa el mismo patrón del módulo común del cap2: variables de entorno + credential.
 */
export function getGraphClient(): Client {
  if (_client) return _client;

  const tenantId = process.env.GRAPH_TENANT_ID;
  const clientId = process.env.GRAPH_CLIENT_ID;
  const clientSecret = process.env.GRAPH_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error(
      'Faltan GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET en el entorno. Copia .env.example a .env.'
    );
  }

  // TokenCredentialProvider que @microsoft/microsoft-graph-client acepta
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

// Tipos mínimos para tipar las respuestas (evita `any` — coherente con cap14 GraphService)
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