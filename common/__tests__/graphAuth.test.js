/**
 * Tests del módulo de auth compartido (JavaScript).
 * No tocan la red: el token se obtiene (con red) solo al hacer la primera petición a Graph.
 * El modo certificado genera un PEM self-signed con openssl en beforeAll (sin deps extra).
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execSync } = require('child_process');
const { getGraphClient, getAccessToken, buildCredential, requireEnv } = require('../graphAuth');

const ENV_KEYS = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET',
                  'CERTIFICATE_PATH', 'CERTIFICATE_PASSWORD'];

// Genera un PEM self-signed válido (clave privada RSA + cert X.509) con openssl al cargar
// el módulo (síncrono, antes de evaluar los describe/it.skip). ClientCertificateCredential
// parsea el PEM al construir (no es lazy para el cert), así que necesitamos un PEM
// criptográficamente válido; el token sí es lazy. Si openssl no está disponible, los tests
// de cert se skiparán (certPem === null).
let tmpDir = null;
let certPem = null;
try {
    execSync('openssl version', { stdio: 'ignore' });
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'lesp-cert-'));
    const keyPem = path.join(tmpDir, 'k.pem');
    const certPemFile = path.join(tmpDir, 'c.pem');
    certPem = path.join(tmpDir, 'fake.pem');
    execSync(
        `openssl req -x509 -newkey rsa:2048 -keyout "${keyPem}" -out "${certPemFile}" ` +
        `-days 1 -nodes -subj "/CN=fake" 2>/dev/null`
    );
    fs.writeFileSync(certPem, fs.readFileSync(keyPem) + fs.readFileSync(certPemFile));
} catch { certPem = null; }

afterAll(() => {
    if (tmpDir) { for (const f of fs.readdirSync(tmpDir)) fs.unlinkSync(path.join(tmpDir, f)); fs.rmdirSync(tmpDir); }
});

beforeEach(() => {
    ENV_KEYS.forEach((k) => delete process.env[k]);
});

function setSecretEnv() {
    process.env.TENANT_ID = 'fake-tenant';
    process.env.CLIENT_ID = 'fake-client';
    process.env.CLIENT_SECRET = 'fake-secret';
}

describe('requireEnv', () => {
    test('throws when env var is missing', () => {
        expect(() => requireEnv('TENANT_ID')).toThrow('TENANT_ID');
    });
});

describe('getGraphClient', () => {
    test('throws when env vars are missing', async () => {
        await expect(getGraphClient()).rejects.toThrow('TENANT_ID');
    });

    test('returns a Graph Client with env set (no network)', async () => {
        setSecretEnv();
        const client = await getGraphClient();
        expect(client).toBeTruthy();
    });
});

// --- Certificado ---

describe('certificate mode', () => {
    const maybeIt = certPem ? it : it.skip;

    maybeIt('auto-detects cert when CERTIFICATE_PATH present (no CLIENT_SECRET needed)', async () => {
        process.env.TENANT_ID = 'fake-tenant';
        process.env.CLIENT_ID = 'fake-client';
        process.env.CERTIFICATE_PATH = certPem;
        // Sin CLIENT_SECRET: debe entrar en modo cert y construir sin lanzar por falta de secret.
        const client = await getGraphClient();
        expect(client).toBeTruthy();
    });

    maybeIt('forces secret when CERTIFICATE_PATH present but useCertificate=false', async () => {
        process.env.TENANT_ID = 't';
        process.env.CLIENT_ID = 'c';
        process.env.CERTIFICATE_PATH = '/some/cert.pfx';
        await expect(getGraphClient({ useCertificate: false })).rejects.toThrow('CLIENT_SECRET');
    });

    test('throws when useCertificate=true but CERTIFICATE_PATH missing', () => {
        process.env.TENANT_ID = 't';
        process.env.CLIENT_ID = 'c';
        expect(() => buildCredential({ useCertificate: true })).toThrow('CERTIFICATE_PATH');
    });

    maybeIt('getAccessToken builds credential in cert mode without network call', async () => {
        // getAccessToken llama a getToken (con red), pero buildCredential por sí no toca red.
        // Verificamos que la construcción del credential en modo cert no lanza por falta de secret.
        process.env.TENANT_ID = 'fake-tenant';
        process.env.CLIENT_ID = 'fake-client';
        process.env.CERTIFICATE_PATH = certPem;
        const credential = buildCredential();
        expect(credential).toBeTruthy();
    });
});