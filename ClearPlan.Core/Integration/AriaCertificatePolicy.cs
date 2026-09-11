using System;
using System.Linq;
using System.Net.Http;
using System.Net.Security;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace ClearPlan.Core.Integration
{
    /// <summary>Normal OS trust by default; optional exact endpoint/certificate binding for configured private installations.</summary>
    public static class AriaCertificatePolicy
    {
        public static HttpClient CreateClient(Uri endpoint,string certificateSha256,int timeoutSeconds)
        {
            if (endpoint==null || endpoint.Scheme!="https" || endpoint.UserInfo.Length!=0)
                throw new ArgumentException("An explicit HTTPS endpoint is required.");
            // Explicit per-client TLS 1.2 also works in legacy ESAPI/standalone hosts whose
            // process defaults predate SystemDefault. Never change the process-wide policy.
            var handler=new HttpClientHandler { AllowAutoRedirect=false, UseCookies=false,
                SslProtocols=System.Security.Authentication.SslProtocols.Tls12 };
            handler.ServerCertificateCustomValidationCallback=(request,certificate,chain,errors)=>
                Validate(endpoint,request.RequestUri,certificateSha256,certificate,errors,
                    chain==null ? null : chain.ChainStatus.Select(s=>s.Status).ToArray(),DateTime.UtcNow);
            return new HttpClient(handler) { Timeout=TimeSpan.FromSeconds(timeoutSeconds) };
        }

        public static bool Validate(Uri expectedEndpoint,Uri requestUri,string expectedSha256,X509Certificate2 certificate,
            SslPolicyErrors errors,X509ChainStatusFlags[] chainErrors,DateTime utcNow)
        {
            if (expectedEndpoint==null || requestUri==null || certificate==null || !requestUri.IsAbsoluteUri ||
                !expectedEndpoint.IsAbsoluteUri || expectedEndpoint.Scheme!="https" || requestUri.Scheme!="https" ||
                requestUri.UserInfo.Length!=0 || !string.Equals(expectedEndpoint.Host,requestUri.Host,StringComparison.OrdinalIgnoreCase) ||
                expectedEndpoint.Port!=requestUri.Port || utcNow.Kind!=DateTimeKind.Utc ||
                utcNow<certificate.NotBefore.ToUniversalTime() || utcNow>certificate.NotAfter.ToUniversalTime()) return false;
            if (string.IsNullOrEmpty(expectedSha256)) return errors==SslPolicyErrors.None;
            string actual;
            using(var sha=SHA256.Create()) actual=BitConverter.ToString(sha.ComputeHash(certificate.RawData)).Replace("-","");
            if (!string.Equals(expectedSha256,actual,StringComparison.OrdinalIgnoreCase)) return false;
            if (errors==SslPolicyErrors.None) return true;
            // Only an untrusted private root may be replaced by explicit leaf trust.
            // Hostname mismatch, expiry, missing certificate, revocation and other chain failures stay fatal.
            return errors==SslPolicyErrors.RemoteCertificateChainErrors && chainErrors!=null && chainErrors.Length>0 &&
                chainErrors.All(error=>error==X509ChainStatusFlags.UntrustedRoot || error==X509ChainStatusFlags.NoError) &&
                chainErrors.Any(error=>error==X509ChainStatusFlags.UntrustedRoot);
        }
    }
}
