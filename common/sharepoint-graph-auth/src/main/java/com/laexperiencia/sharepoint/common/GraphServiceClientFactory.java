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
 * Single factory for creating an authenticated (app-only) {@link GraphServiceClient} from
 * environment variables. Supports two modes:
 *
 * <ul>
 *   <li><b>client secret</b> (default): reads {@code TENANT_ID}/{@code CLIENT_ID}/{@code CLIENT_SECRET}.</li>
 *   <li><b>certificate</b>: reads {@code TENANT_ID}/{@code CLIENT_ID}/{@code CERTIFICATE_PATH}
 *       (+ optional {@code CERTIFICATE_PASSWORD}). Supports {@code .pem} (via
 *       {@code pemCertificate}) and {@code .pfx}/{@code .p12} (via {@code pfxCertificate}).</li>
 * </ul>
 *
 * <p>Use {@link #create()} to auto-detect (cert if {@code CERTIFICATE_PATH} is present,
 * secret otherwise). Mains of chapters 3-7 should call {@code create()}.
 * Replaces the duplicated auth in each Java chapter.
 */
public final class GraphServiceClientFactory {

    private GraphServiceClientFactory() {}

    /**
     * Auto-detects the mode: certificate if {@code CERTIFICATE_PATH} is present and
     * {@code CLIENT_SECRET} is absent; otherwise client secret. Recommended for the
     * mains of chapters 3-7 (same contract as {@code GraphAuthOptions.UsesCertificate} in C#).
     *
     * @return an authenticated {@link GraphServiceClient}.
     */
    public static GraphServiceClient create() {
        String certPath = envOrNone("CERTIFICATE_PATH");
        boolean useCertificate = certPath != null;
        return create(useCertificate);
    }

    /**
     * Builds the client forcing the indicated mode.
     * @param useCertificate {@code true} for certificate, {@code false} for client secret.
     */
    public static GraphServiceClient create(boolean useCertificate) {
        return useCertificate ? createFromCertificate() : createFromSecret();
    }

    /** Reads {@code TENANT_ID}, {@code CLIENT_ID} and {@code CLIENT_SECRET} from the environment and builds the client. */
    public static GraphServiceClient createFromSecret() {
        return createFromSecret(requireEnv("TENANT_ID"), requireEnv("CLIENT_ID"), requireEnv("CLIENT_SECRET"));
    }

    /** Variant with explicit parameters (for tests / dependency injection). */
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
     * Reads {@code TENANT_ID}, {@code CLIENT_ID} and {@code CERTIFICATE_PATH} (+ optional
     * {@code CERTIFICATE_PASSWORD}) from the environment and builds the client with a certificate.
     * Distinguishes {@code .pem} from {@code .pfx}/{@code .p12} by the path extension.
     */
    public static GraphServiceClient createFromCertificate() {
        return createFromCertificate(
                requireEnv("TENANT_ID"),
                requireEnv("CLIENT_ID"),
                requireEnv("CERTIFICATE_PATH"),
                envOrNone("CERTIFICATE_PASSWORD"));
    }

    /** Variant with explicit parameters (for tests / dependency injection). */
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
            // .pfx or .p12
            builder.pfxCertificate(certificatePath, certificatePassword);
        }

        ClientCertificateCredential credential = builder.build();
        return buildClient(credential);
    }

    /**
     * Builds the {@link GraphServiceClient} from any {@link com.azure.core.credential.TokenCredential}.
     *
     * <p>WORKAROUND (bug microsoft-graph 6.1.0 + kiota-auth-azure 1.0.0):
     * {@code GraphServiceClient(credential, String...)} builds the Kiota provider passing an
     * empty (but non-null) varargs array (additionalScopes). Kiota's constructor does:
     * <pre> if (additionalScopes != null)  _scopes = Arrays.asList(additionalScopes);  // immutable
     *  else                            _scopes = new ArrayList&lt;&gt();                 // mutable</pre>
     * and then {@code getAuthorizationToken} calls {@code _scopes.add(...)}. With a non-null empty
     * array, _scopes is immutable and {@code .add()} throws {@code UnsupportedOperationException}.
     * Fix: build the provider by hand passing {@code (String[]) null} to the varargs -> the
     * {@code new ArrayList&lt;&gt();} branch (mutable). We use the core provider (microsoft-graph-core)
     * so national clouds (graph.microsoft.com, etc.) stay configured.
     */
    private static GraphServiceClient buildClient(com.azure.core.credential.TokenCredential credential) {
        AzureIdentityAccessTokenProvider tokenProvider = new AzureIdentityAccessTokenProvider(
                credential,
                (String[]) null,            // scopes (not used; core overrides the host validator)
                (ObservabilityOptions) null, // default observability
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