/*
 * SiteOperations.java
 * Chapter 03: Sites
 *
 * SharePoint Site Operations Example
 * Demonstrates listing, creating, and retrieving SharePoint sites
 */

package com.sharepointexperience.chapter03;

import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.List;
import java.util.Map;
import java.util.HashMap;
import com.google.gson.Gson;
import com.google.gson.JsonObject;
import com.google.gson.JsonArray;

/**
 * Site data class
 */
class Site {
    String id;
    String name;
    String displayName;
    String webUrl;
    String description;
}

/**
 * Response wrapper for site list
 */
class SiteListResponse {
    JsonArray value;
}

/**
 * Example class demonstrating SharePoint site operations
 */
public class SiteOperations {
    private final String tenantId;
    private final String clientId;
    private final String clientSecret;
    private final HttpClient httpClient;
    private final Gson gson;
    private String accessToken;

    /**
     * Initializes a new instance of SiteOperations
     */
    public SiteOperations() {
        this.tenantId = System.getenv("TENANT_ID");
        this.clientId = System.getenv("CLIENT_ID");
        this.clientSecret = System.getenv("CLIENT_SECRET");
        this.httpClient = HttpClient.newHttpClient();
        this.gson = new Gson();

        validateConfig();
    }

    private void validateConfig() {
        if (tenantId == null || clientId == null || clientSecret == null) {
            throw new IllegalArgumentException(
                "TENANT_ID, CLIENT_ID, and CLIENT_SECRET environment variables are required"
            );
        }
    }

    /**
     * Gets access token for Microsoft Graph
     */
    private String getAccessToken() throws IOException, InterruptedException {
        if (accessToken != null) return accessToken;

        String tokenEndpoint = String.format(
            "https://login.microsoftonline.com/%s/oauth2/v2.0/token",
            tenantId
        );

        String form = String.format(
            "client_id=%s\u0026client_secret=%s\u0026scope=%s\u0026grant_type=client_credentials",
            clientId, clientSecret, "https://graph.microsoft.com/.default"
        );

        HttpRequest request = HttpRequest.newBuilder()
            .uri(URI.create(tokenEndpoint))
            .header("Content-Type", "application/x-www-form-urlencoded")
            .POST(HttpRequest.BodyPublishers.ofString(form))
            .build();

        HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());

        if (response.statusCode() != 200) {
            throw new IOException("Failed to get access token: " + response.body());
        }

        JsonObject jsonResponse = gson.fromJson(response.body(), JsonObject.class);
        accessToken = jsonResponse.get("access_token").getAsString();
        return accessToken;
    }

    /**
     * Lists all sites in the organization
     */
    public void listSites() throws IOException, InterruptedException {
        try {
            System.out.println("Fetching all sites...\n");

            String token = getAccessToken();
            HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create("https://graph.microsoft.com/v1.0/sites"))
                .header("Authorization", "Bearer " + token)
                .header("Content-Type", "application/json")
                .GET()
                .build();

            HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());

            if (response.statusCode() != 200) {
                throw new IOException("Failed to list sites: " + response.body());
            }

            JsonObject jsonResponse = gson.fromJson(response.body(), JsonObject.class);
            JsonArray sites = jsonResponse.getAsJsonArray("value");

            System.out.println("Found " + sites.size() + " sites:");
            System.out.println("-".repeat(80));

            for (int i = 0; i < sites.size(); i++) {
                JsonObject site = sites.get(i).getAsJsonObject();
                System.out.println("Display Name: " + safeStr(site, "displayName", safeStr(site, "name", "N/A")));
                System.out.println("  ID: " + safeStr(site, "id", "N/A"));
                System.out.println("  Web URL: " + safeStr(site, "webUrl", "N/A"));
                System.out.println("-".repeat(80));
            }

        } catch (Exception e) {
            System.err.println("Error listing sites: " + e.getMessage());
            throw e;
        }
    }

    /** Returns the string value of a key from a JsonObject, or default if null/absent. */
    private String safeStr(JsonObject obj, String key, String defaultValue) {
        if (obj.get(key) != null && !obj.get(key).isJsonNull()) {
            return obj.get(key).getAsString();
        }
        return defaultValue;
    }

    /**
     * Gets the root site of the organization
     */
    public void getRootSite() throws IOException, InterruptedException {
        try {
            System.out.println("Fetching root site...");

            String token = getAccessToken();
            HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create("https://graph.microsoft.com/v1.0/sites/root"))
                .header("Authorization", "Bearer " + token)
                .header("Content-Type", "application/json")
                .GET()
                .build();

            HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());

            if (response.statusCode() != 200) {
                throw new IOException("Failed to get root site: " + response.body());
            }

            JsonObject site = gson.fromJson(response.body(), JsonObject.class);

            System.out.println("\nRoot site:");
            System.out.println("  Display Name: " + safeStr(site, "displayName", safeStr(site, "name", "N/A")));
            System.out.println("  ID: " + safeStr(site, "id", "N/A"));
            System.out.println("  Web URL: " + safeStr(site, "webUrl", "N/A"));

        } catch (Exception e) {
            System.err.println("Error getting root site: " + e.getMessage());
            throw e;
        }
    }

    /**
     * Gets a specific site by ID
     */
    public void getSiteById(String siteId) throws IOException, InterruptedException {
        try {
            System.out.println("Fetching site by ID: " + siteId);

            String token = getAccessToken();
            HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create("https://graph.microsoft.com/v1.0/sites/" + siteId))
                .header("Authorization", "Bearer " + token)
                .header("Content-Type", "application/json")
                .GET()
                .build();

            HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());

            if (response.statusCode() != 200) {
                throw new IOException("Failed to get site: " + response.body());
            }

            JsonObject site = gson.fromJson(response.body(), JsonObject.class);

            System.out.println("\nSite found:");
            System.out.println("  Name: " + site.get("name").getAsString());
            System.out.println("  Web URL: " + site.get("webUrl").getAsString());

        } catch (Exception e) {
            System.err.println("Error getting site by ID: " + e.getMessage());
            throw e;
        }
    }

    /**
     * Entry point for the example
     */
    public static void main(String[] args) {
        try {
            System.out.println("=== SharePoint Site Operations Example ===\n");

            SiteOperations siteOps = new SiteOperations();

            // Get root site
            siteOps.getRootSite();

            // List all sites
            siteOps.listSites();

            System.out.println("\nSite operations completed successfully!");

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            System.exit(1);
        }
    }
}
