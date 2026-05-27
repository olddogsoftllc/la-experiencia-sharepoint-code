/*
 * AuthWithCertificate.cs
 * Chapter 02: Authentication
 *
 * Certificate-Based Authentication Example
 * Demonstrates secure authentication using X.509 certificates
 */

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Graph;

namespace SharePointExperience.Chapter02
{
    /// <summary>
    /// Example class demonstrating certificate-based authentication to Microsoft Graph
    /// </summary>
    public class AuthWithCertificate
    {
        private readonly string _tenantId;
        private readonly string _clientId;
        private readonly string _certificateThumbprint;
        private readonly string _certificatePath;
        private readonly string _certificatePassword;

        /// <summary>
        /// Initializes a new instance of AuthWithCertificate
        /// </summary>
        public AuthWithCertificate()
        {
            _tenantId = Environment.GetEnvironmentVariable("TENANT_ID")
                ?? throw new ArgumentException("TENANT_ID environment variable is required");
            _clientId = Environment.GetEnvironmentVariable("CLIENT_ID")
                ?? throw new ArgumentException("CLIENT_ID environment variable is required");
            _certificateThumbprint = Environment.GetEnvironmentVariable("CERTIFICATE_THUMBPRINT");
            _certificatePath = Environment.GetEnvironmentVariable("CERTIFICATE_PATH");
            _certificatePassword = Environment.GetEnvironmentVariable("CERTIFICATE_PASSWORD");
        }

        /// <summary>
        /// Loads the X.509 certificate from certificate store or file
        /// </summary>
        /// <returns>X509Certificate2 instance</returns>
        private X509Certificate2 LoadCertificate()
        {
            // Try loading from certificate store by thumbprint
            if (!string.IsNullOrEmpty(_certificateThumbprint))
            {
                using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                var certs = store.Certificates.Find(X509FindType.FindByThumbprint, _certificateThumbprint, false);

                if (certs.Count > 0)
                {
                    Console.WriteLine("Certificate loaded from certificate store");
                    return certs[0];
                }
            }

            // Try loading from file
            if (!string.IsNullOrEmpty(_certificatePath) && File.Exists(_certificatePath))
            {
                var cert = string.IsNullOrEmpty(_certificatePassword)
                    ? new X509Certificate2(_certificatePath)
                    : new X509Certificate2(_certificatePath, _certificatePassword);

                Console.WriteLine("Certificate loaded from file");
                return cert;
            }

            throw new InvalidOperationException("Certificate not found. Please provide CERTIFICATE_THUMBPRINT or CERTIFICATE_PATH");
        }

        /// <summary>
        /// Creates an authenticated GraphServiceClient using certificate credentials
        /// </summary>
        /// <returns>Authenticated GraphServiceClient instance</returns>
        public GraphServiceClient GetAuthenticatedClient()
        {
            try
            {
                var certificate = LoadCertificate();

                // Create certificate credential
                var clientCertificateCredential = new ClientCertificateCredential(
                    _tenantId,
                    _clientId,
                    certificate
                );

                // Create and return the Graph client
                var graphClient = new GraphServiceClient(clientCertificateCredential);

                Console.WriteLine("Successfully authenticated using certificate");
                return graphClient;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Certificate authentication failed: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Validates the certificate properties
        /// </summary>
        public void ValidateCertificate()
        {
            try
            {
                var certificate = LoadCertificate();

                Console.WriteLine("Certificate Details:");
                Console.WriteLine($"  Subject: {certificate.Subject}");
                Console.WriteLine($"  Issuer: {certificate.Issuer}");
                Console.WriteLine($"  Thumbprint: {certificate.Thumbprint}");
                Console.WriteLine($"  Valid From: {certificate.NotBefore}");
                Console.WriteLine($"  Valid Until: {certificate.NotAfter}");
                Console.WriteLine($"  Has Private Key: {certificate.HasPrivateKey}");

                // Check expiration
                if (certificate.NotAfter < DateTime.Now)
                {
                    throw new InvalidOperationException("Certificate has expired");
                }

                if (certificate.NotBefore > DateTime.Now)
                {
                    throw new InvalidOperationException("Certificate is not yet valid");
                }

                Console.WriteLine("Certificate validation passed");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Certificate validation failed: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Entry point for the example
        /// </summary>
        public static void Main(string[] args)
        {
            try
            {
                var authCert = new AuthWithCertificate();

                // Validate certificate first
                authCert.ValidateCertificate();

                // Get authenticated client
                var graphClient = authCert.GetAuthenticatedClient();

                Console.WriteLine("Certificate authentication completed successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Environment.Exit(1);
            }
        }
    }
}
