package com.laexperiencia.sharepoint.common;

import com.azure.identity.ClientCertificateCredential;
import com.azure.identity.ClientCertificateCredentialBuilder;
import com.azure.identity.ClientSecretCredential;
import com.azure.identity.ClientSecretCredentialBuilder;
import com.microsoft.graph.core.authentication.AzureIdentityAccessTokenProvider;
import com.microsoft.graph.serviceclient.GraphServiceClient;
import com.microsoft.kiota.authentication.BaseBearerTokenAuthenticationProvider;
import com.microsoft.kiota.authentication.ObservabilityOptions;

import java.util.Objects;

/**
 * Fábrica única para crear un {@link GraphServiceClient} autenticado (app-only) a partir
 * de variables de entorno. Soporta dos modos:
 *
 * <ul>
 *   <li><b>client secret</b> (por defecto): lee {@code TENANT_ID}/{@code CLIENT_ID}/{@code CLIENT_SECRET}.</li>
 *   <li><b>certificado</b>: lee {@code TENANT_ID}/{@code CLIENT_ID}/{@code CERTIFICATE_PATH}
 *       (+ opcional {@code CERTIFICATE_PASSWORD}). Soporta {@code .pem} (vía
 *       {@code pemCertificate}) y {@code .pfx}/{@code .p12} (vía {@code pfxCertificate}).</li>
 * </ul>
 *
 * <p>Usa {@link #create()} para auto-detectar (cert si {@code CERTIFICATE_PATH} está presente,
 * secret en caso contrario). Los mains de los capítulos 3-7 deben llamar a {@code create()}.
 * Reemplaza la auth duplicada en cada capítulo Java.
 */
public final class GraphServiceClientFactory {

    private GraphServiceClientFactory() {}

    /**
     * Auto-detecta el modo: certificado si {@code CERTIFICATE_PATH} está presente y
     * {@code CLIENT_SECRET} ausente; en caso contrario client secret. Recomendado para los
     * mains de caps 3-7 (mismo contrato que {@code GraphAuthOptions.UsesCertificate} en C#).
     *
     * @return un {@link GraphServiceClient} autenticado.
     */
    public static GraphServiceClient create() {
        String certPath = envOrNone("CERTIFICATE_PATH");
        boolean useCertificate = certPath != null;
        return create(useCertificate);
    }

    /**
     * Construye el cliente forzando el modo indicado.
     * @param useCertificate {@code true} para certificado, {@code false} para client secret.
     */
    public static GraphServiceClient create(boolean useCertificate) {
        return useCertificate ? createFromCertificate() : createFromSecret();
    }

    /** Lee {@code TENANT_ID}, {@code CLIENT_ID} y {@code CLIENT_SECRET} del entorno y construye el cliente. */
    public static GraphServiceClient createFromSecret() {
        return createFromSecret(requireEnv("TENANT_ID"), requireEnv("CLIENT_ID"), requireEnv("CLIENT_SECRET"));
    }

    /** Variante con parámetros explícitos (para tests / inyección de dependencias). */
    public static GraphServiceClient createFromSecret(String tenantId, String clientId, String clientSecret) {
        if (tenantId == null || tenantId.isBlank() || clientId == null || clientId.isBlank()
                || clientSecret == null || clientSecret.isBlank()) {
            throw new IllegalArgumentException("createFromSecret requiere tenantId, clientId y clientSecret no vacíos");
        }
        ClientSecretCredential credential = new ClientSecretCredentialBuilder()
                .tenantId(tenantId)
                .clientId(clientId)
                .clientSecret(clientSecret)
                .build();

        return buildClient(credential);
    }

    /**
     * Lee {@code TENANT_ID}, {@code CLIENT_ID} y {@code CERTIFICATE_PATH} (+ opcional
     * {@code CERTIFICATE_PASSWORD}) del entorno y construye el cliente con certificado.
     * Distingue {@code .pem} de {@code .pfx}/{@code .p12} por la extensión del path.
     */
    public static GraphServiceClient createFromCertificate() {
        return createFromCertificate(
                requireEnv("TENANT_ID"),
                requireEnv("CLIENT_ID"),
                requireEnv("CERTIFICATE_PATH"),
                envOrNone("CERTIFICATE_PASSWORD"));
    }

    /** Variante con parámetros explícitos (para tests / inyección de dependencias). */
    public static GraphServiceClient createFromCertificate(String tenantId, String clientId,
                                                            String certificatePath, String certificatePassword) {
        if (tenantId == null || tenantId.isBlank() || clientId == null || clientId.isBlank()) {
            throw new IllegalArgumentException("createFromCertificate requiere tenantId y clientId no vacíos");
        }
        if (certificatePath == null || certificatePath.isBlank()) {
            throw new IllegalArgumentException("Falta la variable de entorno CERTIFICATE_PATH");
        }

        ClientCertificateCredentialBuilder builder = new ClientCertificateCredentialBuilder()
                .tenantId(tenantId)
                .clientId(clientId);

        String lower = certificatePath.toLowerCase();
        if (lower.endsWith(".pem")) {
            builder.pemCertificate(certificatePath);
        } else {
            // .pfx o .p12
            builder.pfxCertificate(certificatePath, certificatePassword);
        }

        ClientCertificateCredential credential = builder.build();
        return buildClient(credential);
    }

    /**
     * Construye el {@link GraphServiceClient} desde cualquier {@link com.azure.core.credential.TokenCredential}.
     *
     * <p>WORKAROUND (bug microsoft-graph 6.1.0 + kiota-auth-azure 1.0.0):
     * {@code GraphServiceClient(credential, String...)} construye el provider de Kiota pasando un
     * array varargs (additionalScopes) vacío PERO NO-NULL. El constructor de Kiota hace:
     * <pre> if (additionalScopes != null)  _scopes = Arrays.asList(additionalScopes);  // inmutable
     *  else                            _scopes = new ArrayList&lt;&gt();                 // mutable</pre>
     * y luego {@code getAuthorizationToken} invoca {@code _scopes.add(...)}. Con array vacío no-null,
     * _scopes queda inmutable y {@code .add()} lanza {@code UnsupportedOperationException}.
     * Solución: construir el provider a mano pasando {@code (String[]) null} al varargs → rama
     * {@code new ArrayList&lt;&gt;()} (mutable). Usamos el provider del core (microsoft-graph-core)
     * para que queden configurados los hosts nacionales (graph.microsoft.com, etc.).
     */
    private static GraphServiceClient buildClient(com.azure.core.credential.TokenCredential credential) {
        AzureIdentityAccessTokenProvider tokenProvider = new AzureIdentityAccessTokenProvider(
                credential,
                (String[]) null,            // scopes (no se usa; el core sobrescribe el host validator)
                (ObservabilityOptions) null, // observability por defecto
                (String[]) null);            // additionalScopes (varargs) — null => _scopes mutable
        BaseBearerTokenAuthenticationProvider authProvider =
                new BaseBearerTokenAuthenticationProvider(tokenProvider);
        return new GraphServiceClient(authProvider);
    }

    private static String requireEnv(String name) {
        String value = System.getenv(name);
        if (Objects.isNull(value) || value.isBlank()) {
            throw new IllegalArgumentException("Falta la variable de entorno " + name);
        }
        return value;
    }

    private static String envOrNone(String name) {
        String value = System.getenv(name);
        if (Objects.isNull(value) || value.isBlank()) {
            return null;
        }
        return value;
    }
}