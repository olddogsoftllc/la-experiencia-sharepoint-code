/*
 * DocumentOperations.java
 * Chapter 04: Documents
 *
 * SharePoint Document Operations Example
 * Demonstrates upload, download, and search operations for documents
 */

package com.sharepointexperience.chapter04;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;
import java.util.HashMap;
import com.google.gson.Gson;
import com.google.gson.JsonObject;
import com.google.gson.JsonArray;

/**
 * Example class demonstrating SharePoint document operations
 */
public class DocumentOperations {
    private final String tenantId;
    private final String clientId;
    private final String clientSecret;
    private final HttpClient httpClient;
    private final Gson gson;
    private String accessToken;

    /**
     * Initializes a new instance of DocumentOperations
     */
    public DocumentOperations() {
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
            "client_id=%s&client_secret=%s&scope=%s&grant_type=client_credentials",
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
     * Uploads a file to SharePoint
     */
    public void uploadFile(String siteId, String driveId, String filePath, String destinationFileName)
            throws IOException, InterruptedException {
        try {
            System.out.println("Uploading file: " + filePath);

            Path path = Paths.get(filePath);
            if (!Files.exists(path)) {
                throw new IOException("File not found: " + filePath);
            }

            byte[] fileContent = Files.readAllBytes(path);
            String encodedName = java.net.URLEncoder.encode(destinationFileName, "UTF-8");

            String token = getAccessToken();
            String url = String.format(
                "https://graph.microsoft.com/v1.0/sites/%s/drives/%s/items/root:/%s:/content",
                siteId, driveId, encodedName
            );

            HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + token)
                .header("Content-Type", "application/octet-stream")
                .PUT(HttpRequest.BodyPublishers.ofByteArray(fileContent))
                .build();

            HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());

            if (response.statusCode() != 200 && response.statusCode() != 201) {
                throw new IOException("Failed to upload file: " + response.body());
            }

            JsonObject result = gson.fromJson(response.body(), JsonObject.class);
            System.out.println("File uploaded successfully:");
            System.out.println("  ID: " + result.get("id").getAsString());
            System.out.println("  Name: " + result.get("name").getAsString());
            System.out.println("  Size: " + result.get("size").getAsLong() + " bytes");
            System.out.println("  Web URL: " + result.get("webUrl").getAsString());

        } catch (Exception e) {
            System.err.println("Error uploading file: " + e.getMessage());
            throw e;
        }
    }

    /**
     * Downloads a file from SharePoint
     */
    public void downloadFile(String siteId, String driveId, String itemId, String downloadPath)
            throws IOException, InterruptedException {
        try {
            System.out.println("Downloading file to: " + downloadPath);

            String token = getAccessToken();
            String url = String.format(
                "https://graph.microsoft.com/v1.0/sites/%s/drives/%s/items/%s/content",
                siteId, driveId, itemId
            );

            HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + token)
                .GET()
                .build();

            HttpResponse<byte[]> response = httpClient.send(request, HttpResponse.BodyHandlers.ofByteArray());

            if (response.statusCode() != 200) {
                throw new IOException("Failed to download file: " + response.statusCode());
            }

            try (FileOutputStream fos = new FileOutputStream(downloadPath)) {
                fos.write(response.body());
            }

            System.out.println("File downloaded successfully to: " + downloadPath);

        } catch (Exception e) {
            System.err.println("Error downloading file: " + e.getMessage());
            throw e;
        }
    }

    /**
     * Lists files in a specific folder
     */
    public void listFilesInFolder(String siteId, String driveId, String folderPath)
            throws IOException, InterruptedException {
        try {
            System.out.println("Listing files in folder: " + (folderPath.isEmpty() ? "root" : folderPath));

            String token = getAccessToken();
            String url;

            if (folderPath.isEmpty()) {
                url = String.format(
                    "https://graph.microsoft.com/v1.0/sites/%s/drives/%s/root/children",
                    siteId, driveId
                );
            } else {
                String encodedPath = java.net.URLEncoder.encode(folderPath, "UTF-8");
                url = String.format(
                    "https://graph.microsoft.com/v1.0/sites/%s/drives/%s/items/root:/%s:/children",
                    siteId, driveId, encodedPath
                );
            }

            HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + token)
                .header("Content-Type", "application/json")
                .GET()
                .build();

            HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());

            if (response.statusCode() != 200) {
                throw new IOException("Failed to list files: " + response.body());
            }

            JsonObject jsonResponse = gson.fromJson(response.body(), JsonObject.class);
            JsonArray items = jsonResponse.getAsJsonArray("value");

            System.out.println("Found " + items.size() + " items:");
            System.out.println("-".repeat(80));

            for (int i = 0; i < items.size(); i++) {
                JsonObject item = items.get(i).getAsJsonObject();
                String itemType = item.has("folder") ? "Folder" : "File";
                System.out.println(itemType + ": " + item.get("name").getAsString());
                System.out.println("  ID: " + item.get("id").getAsString());
                System.out.println("  Size: " + item.get("size").getAsLong() + " bytes");
                System.out.println("-".repeat(80));
            }

        } catch (Exception e) {
            System.err.println("Error listing files: " + e.getMessage());
            throw e;
        }
    }

    /**
     * Entry point for the example
     */
    public static void main(String[] args) {
        try {
            System.out.println("=== SharePoint Document Operations Example ===\n");

            DocumentOperations docOps = new DocumentOperations();

            System.out.println("Document operations class initialized successfully!");
            System.out.println("Use the methods to perform upload, download, and search operations.");

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            System.exit(1);
        }
    }
}
