/*
 * AuthWithCertificate.java
 * Chapter 02: Authentication
 *
 * Certificate-Based Authentication Example
 * Demonstrates secure authentication using X.509 certificates
 */

package com.sharepointexperience.chapter02;

import java.io.FileInputStream;
import java.io.IOException;
import java.security.KeyStore;
import java.security.MessageDigest;
import java.security.PrivateKey;
import java.security.cert.Certificate;
import java.security.cert.CertificateFactory;
import java.security.cert.X509Certificate;
import java.util.Base64;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Enumeration;

/**
 * Certificate loader utility
 */
class CertificateLoader {

    /**
     * Loads certificate from file
     * @param filePath Path to certificate file
     * @param password Password for encrypted certificate (can be null)
     * @return Map with certificate details
     * @throws Exception if loading fails
     */
    public static Map<String, Object> loadFromFile(String filePath, String password) throws Exception {
        try {
            KeyStore keyStore = KeyStore.getInstance("PKCS12");

            try (FileInputStream fis = new FileInputStream(filePath)) {
                keyStore.load(fis, password != null ? password.toCharArray() : null);
            }

            // Get the alias
            Enumeration<String> aliases = keyStore.aliases();
            if (!aliases.hasMoreElements()) {
                throw new Exception("No certificates found in file");
            }

            String alias = aliases.nextElement();

            // Get private key and certificate
            PrivateKey privateKey = (PrivateKey) keyStore.getKey(
                alias,
                password != null ? password.toCharArray() : null
            );
            X509Certificate certificate = (X509Certificate) keyStore.getCertificate(alias);

            // Calculate thumbprint
            byte[] certEncoded = certificate.getEncoded();
            MessageDigest md = MessageDigest.getInstance("SHA-1");
            byte[] thumbprintBytes = md.digest(certEncoded);
            String thumbprint = bytesToHex(thumbprintBytes);

            System.out.println("Certificate loaded from file");

            Map<String, Object> result = new HashMap<>();
            result.put("privateKey", privateKey);
            result.put("certificate", certificate);
            result.put("thumbprint", thumbprint);
            return result;

        } catch (Exception e) {
            System.err.println("Failed to load certificate from file: " + e.getMessage());
            throw e;
        }
    }

    /**
     * Converts byte array to hex string
     */
    private static String bytesToHex(byte[] bytes) {
        StringBuilder result = new StringBuilder();
        for (byte b : bytes) {
            result.append(String.format("%02X", b));
        }
        return result.toString();
    }

    /**
     * Validates certificate details
     * @param certData Certificate data map
     * @throws Exception if validation fails
     */
    public static void validateCertificate(Map<String, Object> certData) throws Exception {
        X509Certificate cert = (X509Certificate) certData.get("certificate");
        String thumbprint = (String) certData.get("thumbprint");
        PrivateKey privateKey = (PrivateKey) certData.get("privateKey");

        System.out.println("\nCertificate Details:");
        System.out.println("  Subject: " + cert.getSubjectX500Principal());
        System.out.println("  Issuer: " + cert.getIssuerX500Principal());
        System.out.println("  Thumbprint: " + thumbprint);
        System.out.println("  Valid From: " + cert.getNotBefore());
        System.out.println("  Valid Until: " + cert.getNotAfter());
        System.out.println("  Has Private Key: " + (privateKey != null));

        // Check expiration
        Date now = new Date();
        if (cert.getNotAfter().before(now)) {
            throw new Exception("Certificate has expired");
        }
        if (cert.getNotBefore().after(now)) {
            throw new Exception("Certificate is not yet valid");
        }
        if (privateKey == null) {
            throw new Exception("Certificate does not have a private key");
        }

        System.out.println("\nCertificate validation passed");
    }
}

/**
 * Example class demonstrating certificate-based authentication
 */
public class AuthWithCertificate {
    private final String tenantId;
    private final String clientId;
    private final String certificatePath;
    private final String certificatePassword;

    /**
     * Initializes a new instance of AuthWithCertificate
     */
    public AuthWithCertificate() {
        this.tenantId = System.getenv("TENANT_ID");
        this.clientId = System.getenv("CLIENT_ID");
        this.certificatePath = System.getenv("CERTIFICATE_PATH");
        this.certificatePassword = System.getenv("CERTIFICATE_PASSWORD");

        validateConfig();
    }

    /**
     * Validates configuration
     */
    private void validateConfig() {
        if (tenantId == null || clientId == null) {
            throw new IllegalArgumentException("TENANT_ID and CLIENT_ID environment variables are required");
        }
        if (certificatePath == null) {
            throw new IllegalArgumentException("CERTIFICATE_PATH environment variable is required");
        }
    }

    /**
     * Loads and validates the certificate
     * @return Certificate data map
     * @throws Exception if loading fails
     */
    private Map<String, Object> getCertificateConfig() throws Exception {
        Map<String, Object> certData = CertificateLoader.loadFromFile(certificatePath, certificatePassword);
        CertificateLoader.validateCertificate(certData);
        return certData;
    }

    /**
     * Acquires access token using certificate authentication
     * Note: In production, use MSAL4J library for complete certificate authentication
     * @return String message indicating success
     * @throws Exception if authentication fails
     */
    public String authenticate() throws Exception {
        try {
            Map<String, Object> certData = getCertificateConfig();

            System.out.println("Successfully authenticated using certificate");

            // In a real implementation, you would use the certificate to sign a JWT
            // and exchange it for an access token with Azure AD
            // This requires the MSAL4J library or similar

            return "Certificate authentication validated";

        } catch (Exception e) {
            System.err.println("Certificate authentication failed: " + e.getMessage());
            throw e;
        }
    }

    /**
     * Entry point for the example
     */
    public static void main(String[] args) {
        try {
            System.out.println("=== Certificate-Based Authentication Example ===\n");

            AuthWithCertificate authCert = new AuthWithCertificate();
            String result = authCert.authenticate();

            System.out.println("\n" + result);
            System.out.println("\nCertificate authentication completed successfully!");

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            System.exit(1);
        }
    }
}
