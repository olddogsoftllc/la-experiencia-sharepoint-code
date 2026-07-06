/**
 * Autenticación compartida para los ejemplos de La Experiencia SharePoint (JavaScript/Node).
 *
 * Una sola fuente de verdad para obtener un cliente de Microsoft Graph autenticado
 * (app-only) a partir de variables de entorno. Soporta dos modos:
 *
 * - **client secret** (por defecto): lee ``TENANT_ID``/``CLIENT_ID``/``CLIENT_SECRET``.
 * - **certificado**: lee ``TENANT_ID``/``CLIENT_ID``/``CERTIFICATE_PATH`` (+ opcional
 *   ``CERTIFICATE_PASSWORD``). Se usa automáticamente cuando ``CERTIFICATE_PATH``
 *   está presente; se puede forzar con ``useCertificate: true/false``.
 *
 * Reemplaza el token hand-rolled con axios que cada capítulo duplicaba. Las credenciales
 * son lazy (el token se obtiene, con red, solo al hacer la primera petición a Graph),
 * así que se puede construir el cliente sin red.
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
        throw new Error(`Falta la variable de entorno ${name}`);
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
 * Construye el TokenCredential apropiado (secret o certificado).
 * @param {object} [opts]
 * @param {string} [opts.tenantId]
 * @param {string} [opts.clientId]
 * @param {string} [opts.clientSecret]
 * @param {string} [opts.certificatePath]
 * @param {string} [opts.certificatePassword]
 * @param {boolean} [opts.useCertificate] - Si undefined, auto-detecta por CERTIFICATE_PATH.
 * @returns {import('@azure/identity').TokenCredential}
 */
function buildCredential(opts = {}) {
    const tenantId = opts.tenantId || requireEnv('TENANT_ID');
    const clientId = opts.clientId || requireEnv('CLIENT_ID');
    const clientSecret = opts.clientSecret !== undefined ? opts.clientSecret : envOrNone('CLIENT_SECRET');
    const certificatePath = opts.certificatePath !== undefined ? opts.certificatePath : envOrNone('CERTIFICATE_PATH');
    const certificatePassword = opts.certificatePassword !== undefined ? opts.certificatePassword : envOrNone('CERTIFICATE_PASSWORD');

    // Auto-detección: si hay CERTIFICATE_PATH, usar certificado salvo que se force secret.
    let useCertificate = opts.useCertificate;
    if (useCertificate === undefined) {
        useCertificate = certificatePath !== null;
    }

    if (useCertificate) {
        if (!certificatePath) {
            throw new Error(
                'Autenticación con certificado seleccionada pero falta CERTIFICATE_PATH'
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
 * Crea un cliente de Microsoft Graph autenticado con client credentials.
 * El token se obtiene (con red) únicamente al hacer la primera petición a Graph.
 * @param {object} [opts] - Ver buildCredential.
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
 * Obtiene un access token (bearer) para Graph con client credentials.
 * Para los ejemplos que usan REST crudo (axios/fetch) en lugar del cliente del SDK.
 * @param {object} [opts] - Ver buildCredential.
 * @returns {Promise<string>}
 */
async function getAccessToken(opts = {}) {
    const credential = buildCredential(opts);
    const token = await credential.getToken([GRAPH_SCOPE]);
    return token.token;
}

module.exports = { getGraphClient, getAccessToken, requireEnv, buildCredential };