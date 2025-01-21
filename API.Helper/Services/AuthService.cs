using Azure.Core;
using Azure.Identity;
using API.Helper.Interfaces;
using Microsoft.Extensions.Configuration;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;

namespace API.Helper.Services
{
    public class AuthService(IConfiguration configuration) : IAuthService
    {
        private readonly string? _tenantId = configuration["TenantId"];
        private readonly string? _clientId = configuration["ClientId"];
        private readonly string? _clientSecret = configuration["ClientSecret"];
        private readonly string? _certificatePath = configuration["CertificatePath"];
        private readonly string? _certificatePassword = configuration["CertificatePassword"];

        public async Task<string> GetAccessTokenAsync(string resourceUrl)
        {
            string token;

            if (Debugger.IsAttached)
            {
                if (resourceUrl.Contains(".sharepoint.com"))
                {
                    if (_tenantId == null || _clientId == null || _certificatePath == null || _certificatePassword == null)
                        throw new Exception($"TenantId, ClientId, CertificatePath or CertificatePassword is null. TenantId: {_tenantId}, ClientId: {_clientId}, CertificatePath: {_certificatePath}, CertificatePassword length: {_certificatePassword?.Length}");
                    token = await GetAccessTokenWithClientCertificateAsync(resourceUrl, _tenantId, _clientId, _certificatePath, _certificatePassword);
                }
                else
                {
                    if (_tenantId == null || _clientId == null || _clientSecret == null)
                        throw new Exception($"TenantId, ClientId or ClientSecret is null. TenantId: {_tenantId}, ClientId: {_clientId}, ClientSecret length: {_clientSecret?.Length}");
                    token = await GetAccessTokenWithClientSecretAsync(resourceUrl, _tenantId, _clientId, _clientSecret);
                }
            }
            else
                token = await GetAccessTokenWithManagedIdentityAsync(resourceUrl);

            return token;
        }

        private static async Task<string> GetAccessTokenWithClientCertificateAsync(string resourceUrl, string tenantId, string clientId, string certificatePath, string certificatePassword)
        {
            var certificate = new X509Certificate2(certificatePath, certificatePassword);
            return await GetToken(new ClientCertificateCredential(tenantId, clientId, certificate), resourceUrl);
        }

        private static async Task<string> GetAccessTokenWithClientSecretAsync(string resourceUrl, string tenantId, string clientId, string clientSecret)
        {
            return await GetToken(new ClientSecretCredential(tenantId, clientId, clientSecret), resourceUrl);
        }

        private static async Task<string> GetAccessTokenWithManagedIdentityAsync(string resourceUrl)
        {
            return await GetToken(new ManagedIdentityCredential(), resourceUrl);
        }

        private static async Task<string> GetToken(TokenCredential credential, string resourceUrl)
        {
            return (await credential.GetTokenAsync(new TokenRequestContext(scopes: [resourceUrl + "/.default"]) { }, new CancellationToken())).Token;
        }
    }
}