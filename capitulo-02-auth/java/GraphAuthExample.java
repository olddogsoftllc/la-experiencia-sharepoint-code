package capitulo02;

import com.azure.identity.ClientSecretCredential;
import com.azure.identity.ClientSecretCredentialBuilder;
import com.microsoft.graph.serviceclient.GraphServiceClient;

/**
 * Capítulo 2: Autenticación
 * Client Credentials Flow
 */
public class GraphAuthExample {

    public static void main(String[] args) {
        System.out.println("=== Capítulo 2: Autenticación ===\n");

        String tenantId = System.getenv("SP_TENANT_ID");
        String clientId = System.getenv("SP_CLIENT_ID");
        String clientSecret = System.getenv("SP_CLIENT_SECRET");

        if (tenantId == null || clientId == null || clientSecret == null) {
            System.err.println("Error: Faltan variables de entorno");
            return;
        }

        try {
            ClientSecretCredential credential = new ClientSecretCredentialBuilder()
                .tenantId(tenantId)
                .clientId(clientId)
                .clientSecret(clientSecret)
                .build();

            GraphServiceClient client = new GraphServiceClient(credential);
            var org = client.organization().get();

            System.out.println("✓ Conexión exitosa");
            System.out.println("  Tenant: " + org.getValue().get(0).getDisplayName());

        } catch (Exception e) {
            System.err.println("✗ Error: " + e.getMessage());
        }
    }
}
