package com.laexperiencia.sharepoint.common;

import com.microsoft.graph.serviceclient.GraphServiceClient;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Tests for GraphServiceClientFactory. Azure Identity credentials are lazy: they do not
 * touch the network when constructing the GraphServiceClient.
 *
 * <p>For certificate mode, the tests generate a self-signed PEM and P12 with openssl in
 * {@link BeforeAll} (no extra Java deps) and inject them via the overloads with explicit
 * parameters. If openssl is not available, the cert tests are skipped ({@link Assumptions}).
 *
 * <p>Note: {@code System.getenv()} is immutable in Java at runtime, so the cert tests use
 * the {@code createFromCertificate(tenant, client, path, password)} overloads instead of
 * mutating the environment (as the Python/JS tests do with monkeypatch/process.env).
 */
@TestInstance(TestInstance.Lifecycle.PER_CLASS)
class GraphServiceClientFactoryTest {

    private Path tmpDir;
    private Path pemPath;
    private Path p12Path;

    @BeforeAll
    void setupCertFiles() throws Exception {
        try {
            Process version = new ProcessBuilder("openssl", "version").redirectErrorStream(true).start();
            if (version.waitFor() != 0) { return; }
        } catch (IOException e) {
            return; // openssl not available: cert tests will be skipped
        }

        tmpDir = Files.createTempDirectory("lesp-jcert-");
        Path keyPem = tmpDir.resolve("k.pem");
        Path certPem = tmpDir.resolve("c.pem");
        pemPath = tmpDir.resolve("fake.pem");
        p12Path = tmpDir.resolve("fake.p12");

        exec("openssl", "req", "-x509", "-newkey", "rsa:2048",
                "-keyout", keyPem.toString(), "-out", certPem.toString(),
                "-days", "1", "-nodes", "-subj", "/CN=fake");
        // PEM = private key + certificate (what pemCertificate expects).
        Files.writeString(pemPath, Files.readString(keyPem) + Files.readString(certPem));
        // P12 encrypted with password "changeit" (what pfxCertificate expects).
        exec("openssl", "pkcs12", "-export", "-out", p12Path.toString(),
                "-inkey", keyPem.toString(), "-in", certPem.toString(),
                "-passout", "pass:changeit");
    }

    @AfterAll
    void cleanup() throws IOException {
        if (tmpDir == null) return;
        try (var paths = Files.walk(tmpDir)) {
            paths.sorted(Comparator.reverseOrder()).forEach(p -> {
                try { Files.deleteIfExists(p); } catch (IOException ignored) {}
            });
        }
    }

    private static void exec(String... cmd) throws Exception {
        Process p = new ProcessBuilder(cmd).redirectErrorStream(true).start();
        int code = p.waitFor();
        if (code != 0) throw new RuntimeException("Comando fallido: " + String.join(" ", cmd));
    }

    @Test
    void returnsClient_whenSecretEnvPresent() {
        // Fake env injected by surefire (TENANT_ID/CLIENT_ID/CLIENT_SECRET).
        GraphServiceClient client = GraphServiceClientFactory.createFromSecret();
        assertNotNull(client);
    }

    @Test
    void throws_whenEnvMissing() {
        // This test only passes if the process does NOT have TENANT_ID in its environment.
        String tenant = System.getenv("TENANT_ID");
        if (tenant != null && !tenant.isBlank()) return;
        assertThrows(IllegalArgumentException.class, GraphServiceClientFactory::createFromSecret);
    }

    // --- Certificate ---

    @Test
    void createFromCertificate_throws_whenPathNull() {
        assertThrows(IllegalArgumentException.class,
                () -> GraphServiceClientFactory.createFromCertificate("t", "c", null, null));
        assertThrows(IllegalArgumentException.class,
                () -> GraphServiceClientFactory.createFromCertificate("t", "c", "", null));
    }

    @Test
    void createFromCertificate_throws_whenTenantOrClientBlank() {
        assertThrows(IllegalArgumentException.class,
                () -> GraphServiceClientFactory.createFromCertificate("", "c", "/x.p12", null));
        assertThrows(IllegalArgumentException.class,
                () -> GraphServiceClientFactory.createFromCertificate("t", "", "/x.p12", null));
    }

    @Test
    void createFromCertificate_buildsClientFromPem() {
        Assumptions.assumeTrue(pemPath != null, "openssl not available");
        GraphServiceClient client = GraphServiceClientFactory.createFromCertificate(
                "fake-tenant", "fake-client", pemPath.toString(), null);
        assertNotNull(client);
    }

    @Test
    void createFromCertificate_buildsClientFromP12() {
        Assumptions.assumeTrue(p12Path != null, "openssl not available");
        GraphServiceClient client = GraphServiceClientFactory.createFromCertificate(
                "fake-tenant", "fake-client", p12Path.toString(), "changeit");
        assertNotNull(client);
    }

    @Test
    void createFromSecret_withParams_buildsClient() {
        GraphServiceClient client = GraphServiceClientFactory.createFromSecret(
                "fake-tenant", "fake-client", "fake-secret");
        assertNotNull(client);
    }

    @Test
    void createFromSecret_withParams_throws_whenBlank() {
        assertThrows(IllegalArgumentException.class,
                () -> GraphServiceClientFactory.createFromSecret("t", "c", ""));
    }
}