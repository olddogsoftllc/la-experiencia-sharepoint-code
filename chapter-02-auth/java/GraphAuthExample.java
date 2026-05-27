/*
 * GraphAuthExample.java
 * Chapter 02: Authentication
 *
 * Microsoft Graph Client Credentials Authentication Example
 * Demonstrates authenticating to Microsoft Graph using client credentials flow
 */

package com.sharepointexperience.chapter02;

import java.io.IOException;
import java.net.URI;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;
import java.util.stream.Collectors;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

/**
 * Example class demonstrating Microsoft Graph authentication using client credentials
 */
public class GraphAuthExample {
    private final String tenantId;
    private final String clientId;
    private final String clientSecret;
    private final String tokenEndpoint;
    private final String scope;
    private final HttpClient httpClient;

    /**
     * Initializes a new instance of GraphAuthExample
     */
    public GraphAuthExample() {
        this.tenantId = System.getenv("TENANT_ID");
        this.clientId = System.getenv("CLIENT_ID");
        this.clientSecret = System.getenv("CLIENT_SECRET");

        validateConfig();

        this.tokenEndpoint = String.format(
            "https://login.microsoftonline.com/%s/oauth2/v2.0/token",
            this.tenantId
        );
        this.scope = "https://graph.microsoft.com/.default";
        this.httpClient = HttpClient.newHttpClient();
    }

    /**
     * Validates that all required environment variables are present
     */
    private void validateConfig() {
        if (tenantId == null || tenantId.isEmpty()) {
            throw new IllegalArgumentException("TENANT_ID environment variable is required");
        }
        if (clientId == null || clientId.isEmpty()) {
            throw new IllegalArgumentException("CLIENT_ID environment variable is required");
        }
        if (clientSecret == null || clientSecret.isEmpty()) {
            throw new IllegalArgumentException("CLIENT_SECRET environment variable is required");
        }
    }

    /**
     * Acquires an access token using client credentials flow
     * @return Access token string
     * @throws IOException if the request fails
     * @throws InterruptedException if the request is interrupted
     */
    public String getAccessToken() throws IOException, InterruptedException {
        try {
            Map<String, String> parameters = new HashMap<>();
            parameters.put("client_id", clientId);
            parameters.put("client_secret", clientSecret);
            parameters.put("scope", scope);
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
                throw new IOException("Authentication failed: " + response.body());
            }

            JsonObject jsonResponse = JsonParser.parseString(response.body()).getAsJsonObject();
            String accessToken = jsonResponse.get("access_token").getAsString();

            System.out.println("Successfully authenticated to Microsoft Graph");
            return accessToken;

        } catch (Exception e) {
            System.err.println("Authentication failed: " + e.getMessage());
            throw e;
        }
    }

    /**
     * Creates authenticated headers for API requests
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
     * Tests the connection by retrieving organization details
     * @throws IOException if the request fails
     * @throws InterruptedException if the request is interrupted
     */
    public void testConnection() throws IOException, InterruptedException {
        try {
            Map<String, String> headers = getAuthenticatedHeaders();

            HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create("https://graph.microsoft.com/v1.0/organization"))
                .header("Authorization", headers.get("Authorization"))
                .header("Content-Type", "application/json")
                .GET()
                .build();

            HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());

            if (response.statusCode() != 200) {
                throw new IOException("Connection test failed: " + response.body());
            }

            JsonObject jsonResponse = JsonParser.parseString(response.body()).getAsJsonObject();
            JsonObject org = jsonResponse.getAsJsonArray("value").get(0).getAsJsonObject();
            System.out.println("Connected to tenant: " + org.get("displayName").getAsString());

        } catch (Exception e) {
            System.err.println("Connection test failed: " + e.getMessage());
            throw e;
        }
    }

    /**
     * Entry point for the example
     */
    public static void main(String[] args) {
        try {
            System.out.println("=== Microsoft Graph Authentication Example ===");

            GraphAuthExample authExample = new GraphAuthExample();
            authExample.testConnection();

            System.out.println("\nAuthentication completed successfully!");

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            System.exit(1);
        }
    }
}
