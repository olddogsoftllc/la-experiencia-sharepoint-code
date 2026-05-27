/*
 * AuthWithTokenCache.java
 * Chapter 02: Authentication
 *
 * Token Cache Authentication Example
 * Demonstrates persistent token caching for improved performance
 */

package com.sharepointexperience.chapter02;

import java.io.IOException;
import java.net.URI;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.time.Instant;
import java.time.temporal.ChronoUnit;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

/**
 * Token cache entry
 */
class TokenCacheEntry {
    String token;
    Instant expiry;

    TokenCacheEntry(String token, Instant expiry) {
        this.token = token;
        this.expiry = expiry;
    }

    boolean isValid() {
        return Instant.now().isBefore(expiry);
    }
}

/**
 * Example class demonstrating token caching for Microsoft Graph authentication
 */
public class AuthWithTokenCache {
    private final String tenantId;
    private final String clientId;
    private final String clientSecret;
    private final String tokenEndpoint;
    private final HttpClient httpClient;
    private final Map<String, TokenCacheEntry> tokenCache;
    private final int TOKEN_LIFETIME_MINUTES = 55;
    private final String CACHE_KEY = "graph_access_token";

    /**
     * Initializes a new instance of AuthWithTokenCache
     */
    public AuthWithTokenCache() {
        this.tenantId = System.getenv("TENANT_ID");
        this.clientId = System.getenv("CLIENT_ID");
        this.clientSecret = System.getenv("CLIENT_SECRET");

        validateConfig();

        this.tokenEndpoint = String.format(
            "https://login.microsoftonline.com/%s/oauth2/v2.0/token",
            this.tenantId
        );
        this.httpClient = HttpClient.newHttpClient();
        this.tokenCache = new ConcurrentHashMap<>();
    }

    /**
     * Validates configuration
     */
    private void validateConfig() {
        if (tenantId == null || clientId == null || clientSecret == null) {
            throw new IllegalArgumentException(
                "TENANT_ID, CLIENT_ID, and CLIENT_SECRET environment variables are required"
            );
        }
    }

    /**
     * Gets access token from cache or requests new one
     * @return Access token string
     * @throws IOException if the request fails
     * @throws InterruptedException if the request is interrupted
     */
    public String getAccessToken() throws IOException, InterruptedException {
        // Check cache first
        TokenCacheEntry cached = tokenCache.get(CACHE_KEY);
        if (cached != null && cached.isValid()) {
            System.out.println("Using cached access token");
            return cached.token;
        }

        // Request new token
        try {
            Map<String, String> parameters = new HashMap<>();
            parameters.put("client_id", clientId);
            parameters.put("client_secret", clientSecret);
            parameters.put("scope", "https://graph.microsoft.com/.default");
            parameters.put("grant_type", "client_credentials");

            String form = parameters.entrySet()
                .stream()
                .map(e -> e.getKey() + "=" + URLEncoder.encode(e.getValue(), StandardCharsets.UTF_8))
                .collect(Collectors.joining("\u0026"));

            HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(tokenEndpoint))
                .header("Content-Type", "application/x-www-form-urlencoded")
                .POST(HttpRequest.BodyPublishers.ofString(form))
                .build();

            HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());

            if (response.statusCode() != 200) {
                throw new IOException("Token request failed: " + response.body());
            }

            JsonObject jsonResponse = JsonParser.parseString(response.body()).getAsJsonObject();
            String accessToken = jsonResponse.get("access_token").getAsString();

            // Calculate expiry time
            Instant expiry = Instant.now().plus(TOKEN_LIFETIME_MINUTES, ChronoUnit.MINUTES);
            tokenCache.put(CACHE_KEY, new TokenCacheEntry(accessToken, expiry));

            System.out.println("New access token acquired and cached");
            return accessToken;

        } catch (Exception e) {
            System.err.println("Failed to acquire token: " + e.getMessage());
            throw e;
        }
    }

    /**
     * Gets authenticated headers with token caching
     * @return Map of headers
     * @throws IOException if token acquisition fails
     * @throws InterruptedException if token acquisition is interrupted
     */
    public Map<String, String> getAuthenticatedHeaders() throws IOException, InterruptedException {
        String token = getAccessToken();
        Map<String, String> headers = new HashMap<>();
        headers.put("Authorization", "Bearer " + token);
        headers.put("Content-Type", "application/json");
        return headers;
    }

    /**
     * Clears the token cache
     */
    public void clearCache() {
        tokenCache.clear();
        System.out.println("Token cache cleared");
    }

    /**
     * Gets cache status
     * @return Map with cache status information
     */
    public Map<String, Object> getCacheStatus() {
        TokenCacheEntry cached = tokenCache.get(CACHE_KEY);
        Map<String, Object> status = new HashMap<>();
        status.put("hasToken", cached != null);
        status.put("isValid", cached != null && cached.isValid());
        status.put("expiryTime", cached != null ? cached.expiry.toString() : null);
        return status;
    }

    /**
     * Entry point for the example
     */
    public static void main(String[] args) {
        try {
            System.out.println("=== Token Cache Authentication Example ===\n");

            AuthWithTokenCache authCache = new AuthWithTokenCache();

            // First call - will acquire new token
            System.out.println("First call (new token):");
            String token1 = authCache.getAccessToken();
            System.out.println("Token acquired: " + token1.substring(0, 20) + "...");

            // Check cache status
            Map<String, Object> status = authCache.getCacheStatus();
            System.out.println("Cache status: HasToken=" + status.get("hasToken") +
                ", IsValid=" + status.get("isValid") + "\n");

            // Second call - should use cached token
            System.out.println("Second call (from cache):");
            String token2 = authCache.getAccessToken();
            System.out.println("Token from cache: " + token2.substring(0, 20) + "...");

            // Verify tokens match
            System.out.println("\nTokens match: " + token1.equals(token2));

            // Clear cache
            authCache.clearCache();
            System.out.println("After clear - HasToken: " + authCache.getCacheStatus().get("hasToken"));

            System.out.println("\nToken caching example completed!");

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            System.exit(1);
        }
    }
}
