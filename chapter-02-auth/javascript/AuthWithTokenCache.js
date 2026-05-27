/**
 * AuthWithTokenCache.js
 * Chapter 02: Authentication
 *
 * Token Cache Authentication Example
 * Demonstrates persistent token caching for improved performance
 *
 * Required environment variables:
 * - TENANT_ID
 * - CLIENT_ID
 * - CLIENT_SECRET
 */

const axios = require('axios');
const qs = require('querystring');

/**
 * Token cache implementation
 */
class TokenCache {
    constructor() {
        this.cache = new Map();
        this.tokenLifetimeMs = 55 * 60 * 1000; // 55 minutes in milliseconds
    }

    /**
     * Gets a token from cache if valid
     * @param {string} key - Cache key
     * @returns {string|null} Cached token or null
     */
    get(key) {
        const cached = this.cache.get(key);
        if (!cached) return null;

        if (Date.now() < cached.expiry) {
            console.log('Using cached access token');
            return cached.token;
        }

        // Token expired, remove from cache
        this.cache.delete(key);
        return null;
    }

    /**
     * Stores a token in cache
     * @param {string} key - Cache key
     * @param {string} token - Access token
     */
    set(key, token) {
        this.cache.set(key, {
            token,
            expiry: Date.now() + this.tokenLifetimeMs
        });
        console.log('New access token acquired and cached');
    }

    /**
     * Clears the token cache
     */
    clear() {
        this.cache.clear();
        console.log('Token cache cleared');
    }

    /**
     * Gets cache status
     * @returns {Object} Cache status information
     */
    getStatus(key) {
        const cached = this.cache.get(key);
        return {
            hasToken: !!cached,
            isValid: cached ? Date.now() < cached.expiry : false,
            expiryTime: cached ? new Date(cached.expiry).toISOString() : null
        };
    }
}

/**
 * Authentication with token caching
 */
class AuthWithTokenCache {
    constructor() {
        this.tenantId = process.env.TENANT_ID;
        this.clientId = process.env.CLIENT_ID;
        this.clientSecret = process.env.CLIENT_SECRET;
        this.cache = new TokenCache();
        this.cacheKey = 'graph_access_token';

        this.validateConfig();
    }

    validateConfig() {
        const required = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET'];
        const missing = required.filter(key => !process.env[key]);

        if (missing.length > 0) {
            throw new Error(`Missing required environment variables: ${missing.join(', ')}`);
        }
    }

    /**
     * Gets token from cache or requests new one
     * @returns {Promise<string>} Access token
     */
    async getAccessToken() {
        // Check cache first
        const cachedToken = this.cache.get(this.cacheKey);
        if (cachedToken) {
            return cachedToken;
        }

        // Request new token
        try {
            const tokenEndpoint = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
            const requestBody = {
                client_id: this.clientId,
                client_secret: this.clientSecret,
                scope: 'https://graph.microsoft.com/.default',
                grant_type: 'client_credentials'
            };

            const response = await axios.post(
                tokenEndpoint,
                qs.stringify(requestBody),
                { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
            );

            const token = response.data.access_token;
            this.cache.set(this.cacheKey, token);

            return token;
        } catch (error) {
            console.error('Failed to acquire token:', error.response?.data?.error_description || error.message);
            throw error;
        }
    }

    /**
     * Gets authenticated headers with token caching
     * @returns {Promise<Object>} Headers object
     */
    async getAuthenticatedHeaders() {
        const token = await this.getAccessToken();
        return {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        };
    }

    /**
     * Clears the token cache
     */
    clearCache() {
        this.cache.clear();
    }

    /**
     * Gets cache status
     * @returns {Object} Cache status
     */
    getCacheStatus() {
        return this.cache.getStatus(this.cacheKey);
    }
}

/**
 * Main execution function
 */
async function main() {
    try {
        console.log('=== Token Cache Authentication Example ===\n');

        const auth = new AuthWithTokenCache();

        // First call - will acquire new token
        console.log('First call (new token):');
        const token1 = await auth.getAccessToken();
        console.log(`Token acquired: ${token1.substring(0, 20)}...`);

        // Check cache status
        const status = auth.getCacheStatus();
        console.log(`Cache status: HasToken=${status.hasToken}, IsValid=${status.isValid}\n`);

        // Second call - should use cached token
        console.log('Second call (from cache):');
        const token2 = await auth.getAccessToken();
        console.log(`Token from cache: ${token2.substring(0, 20)}...`);

        // Verify tokens match
        console.log(`\nTokens match: ${token1 === token2}`);

        // Clear cache
        auth.clearCache();
        console.log(`After clear - HasToken: ${auth.getCacheStatus().hasToken}`);

        console.log('\nToken caching example completed!');
    } catch (error) {
        console.error('Error:', error.message);
        process.exit(1);
    }
}

module.exports = { AuthWithTokenCache, TokenCache };

if (require.main === module) {
    main();
}
