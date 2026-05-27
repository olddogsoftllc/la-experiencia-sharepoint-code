/**
 * AuthWithCertificate.js
 * Chapter 02: Authentication
 *
 * Certificate-Based Authentication Example
 * Demonstrates secure authentication using X.509 certificates
 *
 * Required environment variables:
 * - TENANT_ID
 * - CLIENT_ID
 * - CERTIFICATE_PATH (or CERTIFICATE_THUMBPRINT with certificate store access)
 * - CERTIFICATE_PASSWORD (optional, for encrypted certificates)
 */

const fs = require('fs');
const crypto = require('crypto');
const { ConfidentialClientApplication } = require('@azure/msal-node');

/**
 * Certificate loader utility
 */
class CertificateLoader {
    /**
     * Loads certificate from file
     * @param {string} filePath - Path to certificate file
     * @param {string} password - Optional password for encrypted certificate
     * @returns {Object} Certificate object with privateKey and thumbprint
     */
    static loadFromFile(filePath, password = null) {
        try {
            if (!fs.existsSync(filePath)) {
                throw new Error(`Certificate file not found: ${filePath}`);
            }

            const certificate = fs.readFileSync(filePath);
            const cert = password
                ? crypto.createPrivateKey({ key: certificate, passphrase: password })
                : crypto.createPrivateKey(certificate);

            // Calculate thumbprint
            const x509Cert = new crypto.X509Certificate(certificate);
            const thumbprint = crypto
                .createHash('sha1')
                .update(x509Cert.raw)
                .digest('hex')
                .toUpperCase();

            console.log('Certificate loaded from file');

            return {
                privateKey: cert.export({ type: 'pkcs1', format: 'pem' }),
                thumbprint: thumbprint
            };
        } catch (error) {
            console.error('Failed to load certificate from file:', error.message);
            throw error;
        }
    }

    /**
     * Validates certificate details
     * @param {Object} certificate - Certificate object
     */
    static validateCertificate(certificate) {
        console.log('\nCertificate Details:');
        console.log(`  Thumbprint: ${certificate.thumbprint}`);
        console.log(`  Has Private Key: ${!!certificate.privateKey}`);

        if (!certificate.privateKey) {
            throw new Error('Certificate does not have a private key');
        }

        console.log('\nCertificate validation passed');
    }
}

/**
 * Certificate-based authentication handler
 */
class AuthWithCertificate {
    constructor() {
        this.tenantId = process.env.TENANT_ID;
        this.clientId = process.env.CLIENT_ID;
        this.certificatePath = process.env.CERTIFICATE_PATH;
        this.certificateThumbprint = process.env.CERTIFICATE_THUMBPRINT;
        this.certificatePassword = process.env.CERTIFICATE_PASSWORD;

        this.validateConfig();
    }

    validateConfig() {
        if (!this.tenantId || !this.clientId) {
            throw new Error('TENANT_ID and CLIENT_ID environment variables are required');
        }

        if (!this.certificatePath && !this.certificateThumbprint) {
            throw new Error('Please provide CERTIFICATE_PATH or CERTIFICATE_THUMBPRINT environment variable');
        }
    }

    /**
     * Gets certificate configuration for MSAL
     * @returns {Object} Certificate configuration
     */
    getCertificateConfig() {
        if (this.certificatePath) {
            const certData = CertificateLoader.loadFromFile(
                this.certificatePath,
                this.certificatePassword
            );
            CertificateLoader.validateCertificate(certData);

            return {
                privateKey: certData.privateKey,
                thumbprint: certData.thumbprint
            };
        }

        // For certificate thumbprint, we'd typically use key vault or certificate store
        throw new Error('Certificate thumbprint authentication requires certificate store access');
    }

    /**
     * Acquires access token using certificate authentication
     * @returns {Promise<string>} Access token
     */
    async getAccessToken() {
        try {
            const certConfig = this.getCertificateConfig();

            const msalConfig = {
                auth: {
                    clientId: this.clientId,
                    authority: `https://login.microsoftonline.com/${this.tenantId}`,
                    clientCertificate: certConfig
                }
            };

            const cca = new ConfidentialClientApplication(msalConfig);

            const tokenRequest = {
                scopes: ['https://graph.microsoft.com/.default']
            };

            const response = await cca.acquireTokenByClientCredential(tokenRequest);

            console.log('Successfully authenticated using certificate');
            return response.accessToken;
        } catch (error) {
            console.error('Certificate authentication failed:', error.message);
            throw error;
        }
    }

    /**
     * Gets authenticated headers
     * @returns {Promise<Object>} Headers object
     */
    async getAuthenticatedHeaders() {
        const token = await this.getAccessToken();
        return {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        };
    }
}

/**
 * Main execution function
 */
async function main() {
    try {
        console.log('=== Certificate-Based Authentication Example ===\n');

        const auth = new AuthWithCertificate();
        const token = await auth.getAccessToken();

        console.log(`\nToken acquired: ${token.substring(0, 30)}...`);
        console.log('\nCertificate authentication completed successfully!');
    } catch (error) {
        console.error('Error:', error.message);
        process.exit(1);
    }
}

module.exports = { AuthWithCertificate, CertificateLoader };

if (require.main === module) {
    main();
}
