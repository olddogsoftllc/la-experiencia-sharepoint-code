/**
 * Tests for the shared auth module (JavaScript).
 * They do not touch the network: the token is fetched (with network) only when making the first request to Graph.
 * Certificate mode generates a self-signed PEM with openssl in beforeAll (no extra deps).
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execSync } = require('child_process');
const { getGraphClient, getAccessToken, buildCredential, requireEnv } = require('../graphAuth');

const ENV_KEYS = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET',
                  'CERTIFICATE_PATH', 'CERTIFICATE_PASSWORD'];

// Generates a valid self-signed PEM (RSA private key + X.509 cert) with openssl when
// loading the module (synchronous, before evaluating describe/it.skip).
// ClientCertificateCredential parses the PEM on construction (it is not lazy for the
// cert), so we need a cryptographically valid PEM; the token itself is lazy. If openssl
// is not available, the cert tests are skipped (certPem === null).
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

// --- Certificate ---

describe('certificate mode', () => {
    const maybeIt = certPem ? it : it.skip;

    maybeIt('auto-detects cert when CERTIFICATE_PATH present (no CLIENT_SECRET needed)', async () => {
        process.env.TENANT_ID = 'fake-tenant';
        process.env.CLIENT_ID = 'fake-client';
        process.env.CERTIFICATE_PATH = certPem;
        // No CLIENT_SECRET: must enter cert mode and build without throwing for missing secret.
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
        // getAccessToken calls getToken (with network), but buildCredential itself does not touch the network.
        // We verify that building the credential in cert mode does not throw for missing secret.
        process.env.TENANT_ID = 'fake-tenant';
        process.env.CLIENT_ID = 'fake-client';
        process.env.CERTIFICATE_PATH = certPem;
        const credential = buildCredential();
        expect(credential).toBeTruthy();
    });
});