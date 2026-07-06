/**
 * Shared authentication for the La Experiencia SharePoint examples (JavaScript/Node).
 *
 * A single source of truth for obtaining an authenticated (app-only) Microsoft Graph
 * client from environment variables. Supports two modes:
 *
 * - **client secret** (default): reads ``TENANT_ID``/``CLIENT_ID``/``CLIENT_SECRET``.
 * - **certificate**: reads ``TENANT_ID``/``CLIENT_ID``/``CERTIFICATE_PATH`` (+ optional
 *   ``CERTIFICATE_PASSWORD``). Used automatically when ``CERTIFICATE_PATH`` is present;
 *   can be forced with ``useCertificate: true/false``.
 *
 * Replaces the hand-rolled token-with-axios that each chapter duplicated. Credentials
 * are lazy (the token is fetched, with network, only when making the first request to
 * Graph), so the client can be built without network access.
 */
const { ClientSecretCredential, ClientCertificateCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');

const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

/**
 * @param {string} name
 * @returns {string}
 */
function requireEnv(name) {
    const value = process.env[name];
    if (!value || !value.trim()) {
        throw new Error(`Missing environment variable ${name}`);
    }
    return value;
}

/**
 * @param {string} name
 * @returns {string|null}
 */
function envOrNone(name) {
    const value = process.env[name];
    if (!value || !value.trim()) {
        return null;
    }
    return value;
}

/**
 * Builds the appropriate TokenCredential (secret or certificate).
 * @param {object} [opts]
 * @param {string} [opts.tenantId]
 * @param {string} [opts.clientId]
 * @param {string} [opts.clientSecret]
 * @param {string} [opts.certificatePath]
 * @param {string} [opts.certificatePassword]
 * @param {boolean} [opts.useCertificate] - If undefined, auto-detected from CERTIFICATE_PATH.
 * @returns {import('@azure/identity').TokenCredential}
 */
function buildCredential(opts = {}) {
    const tenantId = opts.tenantId || requireEnv('TENANT_ID');
    const clientId = opts.clientId || requireEnv('CLIENT_ID');
    const clientSecret = opts.clientSecret !== undefined ? opts.clientSecret : envOrNone('CLIENT_SECRET');
    const certificatePath = opts.certificatePath !== undefined ? opts.certificatePath : envOrNone('CERTIFICATE_PATH');
    const certificatePassword = opts.certificatePassword !== undefined ? opts.certificatePassword : envOrNone('CERTIFICATE_PASSWORD');

    // Auto-detection: if CERTIFICATE_PATH is present, use certificate unless secret is forced.
    let useCertificate = opts.useCertificate;
    if (useCertificate === undefined) {
        useCertificate = certificatePath !== null;
    }

    if (useCertificate) {
        if (!certificatePath) {
            throw new Error(
                'Certificate authentication selected but CERTIFICATE_PATH is missing'
            );
        }
        return new ClientCertificateCredential(
            tenantId,
            clientId,
            {
                certificatePath,
                certificatePassword: certificatePassword || undefined,
            }
        );
    }

    if (!clientSecret) {
        throw new Error(
            'Falta CLIENT_SECRET (o define CERTIFICATE_PATH para usar certificado)'
        );
    }
    return new ClientSecretCredential(tenantId, clientId, clientSecret);
}

/**
 * Creates an authenticated Microsoft Graph client using client credentials.
 * The token is fetched (with network) only when making the first request to Graph.
 * @param {object} [opts] - See buildCredential.
 * @returns {Promise<import('@microsoft/microsoft-graph-client').Client>}
 */
async function getGraphClient(opts = {}) {
    const credential = buildCredential(opts);
    return Client.init({
        authProvider: {
            getAccessToken: async () => {
                const token = await credential.getToken([GRAPH_SCOPE]);
                return token.token;
            },
        },
    });
}

/**
 * Obtains a (bearer) access token for Graph using client credentials.
 * For examples that use raw REST (axios/fetch) instead of the SDK client.
 * @param {object} [opts] - See buildCredential.
 * @returns {Promise<string>}
 */
async function getAccessToken(opts = {}) {
    const credential = buildCredential(opts);
    const token = await credential.getToken([GRAPH_SCOPE]);
    return token.token;
}

module.exports = { getGraphClient, getAccessToken, requireEnv, buildCredential };