package com.olddogsoft.sharepoint.automation;

import java.net.URI;
import java.util.*;

/**
 * DeltaQueryManager.java
 * Ejemplo de Delta Queries con Microsoft Graph para sincronizacion incremental.
 * Referencia: Capitulo 7 - Automatizacion y Flujos
 *
 * Requiere: java.net.http.HttpClient, org.json (o Gson)
 */
public class DeltaQueryManager {

    private final String accessToken;
    private final String baseUrl = "https://graph.microsoft.com/v1.0";
    private final Map<String, String> deltaTokens = new HashMap<>();

    public DeltaQueryManager(String accessToken) {
        if (accessToken == null || accessToken.isBlank()) {
            throw new IllegalArgumentException("accessToken is required");
        }
        this.accessToken = accessToken;
    }

    /**
     * Realiza consulta delta inicial y almacena token.
     */
    public List<Map<String, Object>> getInitialDelta(String driveId, String folderId) throws Exception {
        if (folderId == null) folderId = "root";

        System.out.println("🔄 Ejecutando consulta delta inicial en drive: " + driveId);

        String url = baseUrl + "/drives/" + driveId + "/items/" + folderId + "/delta";
        Map<String, Object> response = executeGet(url);

        @SuppressWarnings("unchecked")
        List<Map<String, Object>> items = (List<Map<String, Object>>) response.getOrDefault("value", new ArrayList<>());
        System.out.println("   Items obtenidos: " + items.size());

        String deltaLink = (String) response.get("@odata.deltaLink");
        String token = extractDeltaToken(deltaLink);
        if (token != null) {
            String key = driveId + "_" + folderId;
            deltaTokens.put(key, token);
            System.out.println("   Token delta guardado: " + token.substring(0, Math.min(50, token.length())) + "...");
        }

        return items;
    }

    /**
     * Sincroniza cambios usando token delta almacenado.
     */
    public List<Map<String, Object>> getDeltaChanges(String driveId, String folderId) throws Exception {
        if (folderId == null) folderId = "root";

        String key = driveId + "_" + folderId;
        if (!deltaTokens.containsKey(key)) {
            System.out.println("⚠️  No existe token delta. Ejecutando consulta inicial...");
            return getInitialDelta(driveId, folderId);
        }

        String token = deltaTokens.get(key);
        System.out.println("🔄 Sincronizando cambios con token delta...");

        String url = baseUrl + "/drives/" + driveId + "/items/" + folderId + "/delta?token=" + token;
        Map<String, Object> response = executeGet(url);

        @SuppressWarnings("unchecked")
        List<Map<String, Object>> items = (List<Map<String, Object>>) response.getOrDefault("value", new ArrayList<>());
        System.out.println("   Cambios detectados: " + items.size());

        for (Map<String, Object> item : items) {
            String name = (String) item.getOrDefault("name", "unknown");
            if (item.containsKey("deleted")) {
                System.out.println("   🗑️  Eliminado: " + name);
            } else {
                System.out.println("   ✏️  Modificado: " + name);
            }
        }

        String newDeltaLink = (String) response.get("@odata.deltaLink");
        String newToken = extractDeltaToken(newDeltaLink);
        if (newToken != null) {
            deltaTokens.put(key, newToken);
            System.out.println("   Token delta actualizado");
        }

        return items;
    }

    /**
     * Consulta delta para items de una lista de SharePoint.
     */
    public List<Map<String, Object>> getListDelta(String siteId, String listId) throws Exception {
        System.out.println("🔄 Consulta delta para lista: " + listId);

        String url = baseUrl + "/sites/" + siteId + "/lists/" + listId + "/items/delta";
        Map<String, Object> response = executeGet(url);

        @SuppressWarnings("unchecked")
        List<Map<String, Object>> items = (List<Map<String, Object>>) response.getOrDefault("value", new ArrayList<>());

        String deltaLink = (String) response.get("@odata.deltaLink");
        String token = extractDeltaToken(deltaLink);
        if (token != null) {
            deltaTokens.put("list_" + siteId + "_" + listId, token);
        }

        return items;
    }

    /**
     * Muestra los tokens delta almacenados.
     */
    public void showStoredTokens() {
        System.out.println("\n📋 Tokens Delta Almacenados:");
        for (Map.Entry<String, String> entry : deltaTokens.entrySet()) {
            String token = entry.getValue();
            System.out.println("   " + entry.getKey() + ": " + token.substring(0, Math.min(50, token.length())) + "...");
        }
    }

    // ============== Metodos Auxiliares ==============

    private Map<String, Object> executeGet(String url) throws Exception {
        // En un ejemplo real, usarias java.net.http.HttpClient
        // Este stub demuestra la estructura
        System.out.println("   GET: " + url);
        return new HashMap<>();
    }

    private String extractDeltaToken(String deltaLink) {
        if (deltaLink == null || deltaLink.isEmpty()) return null;

        try {
            URI uri = new URI(deltaLink);
            String query = uri.getQuery();
            if (query == null) return null;

            for (String param : query.split("&")) {
                String[] parts = param.split("=");
                if (parts.length == 2 && (parts[0].equals("token") || parts[0].equals("deltaToken"))) {
                    return parts[1];
                }
            }
        } catch (Exception e) {
            System.err.println("Error parsing delta link: " + e.getMessage());
        }
        return null;
    }
}
