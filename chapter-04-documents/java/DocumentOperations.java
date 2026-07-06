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
     * Lists the document libraries (drives) of a site. Solo lectura.
     */
    public void listDrives(String siteId) throws IOException, InterruptedException {
        try {
            System.out.println("Listing document libraries of site: " + siteId);
            String token = getAccessToken();
            String url = String.format("https://graph.microsoft.com/v1.0/sites/%s/drives", siteId);

            HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Authorization", "Bearer " + token)
                .header("Content-Type", "application/json")
                .GET()
                .build();

            HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());
            if (response.statusCode() != 200) {
                throw new IOException("Error listing drives: " + response.body());
            }

            JsonObject json = gson.fromJson(response.body(), JsonObject.class);
            var drives = gson.fromJson(json.get("value"), com.google.gson.JsonArray.class);
            System.out.println("Found " + drives.size() + " libraries:");
            System.out.println("-".repeat(80));
            for (var d : drives) {
                JsonObject drive = d.getAsJsonObject();
                System.out.println("Library: " + drive.get("name"));
                System.out.println("  ID: " + drive.get("id"));
                System.out.println("  Type: " + drive.get("driveType"));
                System.out.println("-".repeat(80));
            }
        } catch (Exception e) {
            System.err.println("Error listing libraries: " + e.getMessage());
            throw e;
        }
    }

    /**
     * Entry point: demo de solo lectura contra el sitio de pruebas book-test.
     */
    public static void main(String[] args) {
        try {
            System.out.println("=== SharePoint Document Operations Example ===\n");
            DocumentOperations docOps = new DocumentOperations();

            String hostname = System.getenv("SHAREPOINT_HOSTNAME");
            if (hostname == null || hostname.isBlank()) hostname = "olddogsoft1.sharepoint.com";
            String sitePath = System.getenv("SHAREPOINT_SITE_PATH");
            if (sitePath == null || sitePath.isBlank()) sitePath = "book-test";

            // Resolver el sitio por path para obtener su ID.
            String token = docOps.getAccessToken();
            String siteUrl = String.format(
                "https://graph.microsoft.com/v1.0/sites/%s:/sites/%s",
                java.net.URLEncoder.encode(hostname, "UTF-8"),
                java.net.URLEncoder.encode(sitePath, "UTF-8"));
            HttpRequest siteReq = HttpRequest.newBuilder()
                .uri(URI.create(siteUrl))
                .header("Authorization", "Bearer " + token)
                .GET()
                .build();
            HttpResponse<String> siteResp = docOps.httpClient.send(siteReq, HttpResponse.BodyHandlers.ofString());
            if (siteResp.statusCode() != 200) {
                throw new IOException("Error resolving site: " + siteResp.body());
            }
            String siteId = docOps.gson.fromJson(siteResp.body(), JsonObject.class).get("id").getAsString();

            docOps.listDrives(siteId);
            System.out.println("\nDocument operations completed successfully!");

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            System.exit(1);
        }
    }
}
