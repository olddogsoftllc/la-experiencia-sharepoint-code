// TermStoreExplorer.java
// Example: Explore the Term Store with Microsoft Graph SDK for Java (v6).
// The GraphServiceClient is injected via constructor (DI) from the common module.

package com.sharepointexperiencia.managedmetadata;

import com.laexperiencia.sharepoint.common.GraphServiceClientFactory;
import com.microsoft.graph.serviceclient.GraphServiceClient;

public class TermStoreExplorer {
    private final GraphServiceClient graphClient;

    public TermStoreExplorer(GraphServiceClient graphClient) {
        this.graphClient = graphClient;
    }

    public void explore() {
        System.out.println("=".repeat(60));
        System.out.println("Explorando Term Store");
        System.out.println("=".repeat(60));
        System.out.println();

        try {
            // Get Term Store (API v6: sites().bySiteId(...).termStore())
            var termStore = graphClient.sites().bySiteId("root").termStore().get();

            System.out.println("📚 Term Store Info:");
            System.out.println("   ID: " + termStore.getId());
            System.out.println("   Default Language: " + termStore.getDefaultLanguageTag());
            System.out.println();

            // List groups
            var groups = graphClient.sites().bySiteId("root").termStore().groups().get();

            int groupCount = groups != null && groups.getValue() != null ? groups.getValue().size() : 0;
            System.out.println("📂 Grupos encontrados: " + groupCount + "\n");

            if (groups != null && groups.getValue() != null) {
                for (var group : groups.getValue()) {
                    System.out.println("   📁 " + group.getDisplayName());
                    System.out.println("      ID: " + group.getId());
                    String desc = group.getDescription() != null ? group.getDescription() : "N/A";
                    System.out.println("      Description: " + desc);

                    // Get term sets
                    var sets = graphClient.sites().bySiteId("root")
                            .termStore().groups().byGroupId(group.getId()).sets().get();

                    if (sets != null && sets.getValue() != null) {
                        for (var set : sets.getValue()) {
                            String setName = (set.getLocalizedNames() != null && !set.getLocalizedNames().isEmpty())
                                    ? set.getLocalizedNames().get(0).getName()
                                    : set.getId();
                            System.out.println("         📚 " + setName);
                        }
                    }
                    System.out.println();
                }
            }
        } catch (Exception e) {
            System.err.println("❌ Error: " + e.getMessage());
        }
    }

    public static void main(String[] args) {
        try {
            // create() auto-detects a certificate (if CERTIFICATE_PATH is present) or a client secret.
            var client = GraphServiceClientFactory.create();
            new TermStoreExplorer(client).explore();
            System.out.println("\nTerm Store exploration completed.");
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            System.exit(1);
        }
    }
}