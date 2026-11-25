using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ClearPlan.Core.Integration
{
    /// <summary>Site configuration only. Credentials are never JSON properties or versioned content.</summary>
    public sealed class AriaUploadConfiguration
    {
        public int SchemaVersion { get; set; } = 1;
        public bool Enabled { get; set; }
        public string BaseUrl { get; set; }
        public string TokenUrl { get; set; }
        public string ClientId { get; set; }
        public string Scope { get; set; } = "system/DocumentReference.cruds system/Patient.rs system/Organization.rs system/ValueSet.rs system/Practitioner.rs";
        public string CredentialEnvironmentVariable { get; set; } = "ARIA_FHIR_CLIENT_SECRET";
        public string CredentialEnvFilePath { get; set; }
        public string ProviderReference { get; set; }
        public string PatientIdentifierSystem { get; set; }
        public string DocumentTypeCode { get; set; }
        public string DocumentTypeDisplay { get; set; }
        public string CategoryCode { get; set; } = "PDF";
        public string FhirCertificateSha256 { get; set; }
        public string TokenCertificateSha256 { get; set; }
        public int TimeoutSeconds { get; set; } = 45;
        public int DocumentDateSafetyMinutes { get; set; }
        public string[] AttachmentReadbackRoots { get; set; } = new string[0];

        public static AriaUploadConfiguration Parse(string json)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(json) || Encoding.UTF8.GetByteCount(json)>32768) throw new FormatException();
                var value=JsonConvert.DeserializeObject<AriaUploadConfiguration>(json,new JsonSerializerSettings {
                    TypeNameHandling=TypeNameHandling.None, MissingMemberHandling=MissingMemberHandling.Error, MaxDepth=8 });
                if (value==null || value.SchemaVersion!=1 || value.TimeoutSeconds<5 || value.TimeoutSeconds>120 ||
                    value.DocumentDateSafetyMinutes<0 || value.DocumentDateSafetyMinutes>60 ||
                    !ValidPin(value.FhirCertificateSha256) || !ValidPin(value.TokenCertificateSha256) ||
                    (value.CategoryCode!="PDF" && value.CategoryCode!="TIF")) throw new FormatException();
                if (!OptionalText(value.CredentialEnvFilePath,1024) ||
                    (!string.IsNullOrEmpty(value.CredentialEnvFilePath) && (value.CredentialEnvFilePath.Contains("://") ||
                        value.CredentialEnvFilePath.StartsWith("file:",StringComparison.OrdinalIgnoreCase))) || !OptionalText(value.ClientId,256) ||
                    !Regex.IsMatch(value.CredentialEnvironmentVariable ?? "", "^[A-Za-z_][A-Za-z0-9_]{0,127}$") ||
                    !OptionalText(value.Scope,1024) || string.IsNullOrWhiteSpace(value.Scope) ||
                    !OptionalText(value.DocumentTypeCode,128) || !OptionalText(value.DocumentTypeDisplay,256) ||
                    !OptionalText(value.PatientIdentifierSystem,512) || !OptionalText(value.ProviderReference,128)) throw new FormatException();
                if (!string.IsNullOrEmpty(value.BaseUrl)) ValidateHttps(value.BaseUrl,true);
                if (!string.IsNullOrEmpty(value.TokenUrl)) ValidateHttps(value.TokenUrl,false);
                if (!string.IsNullOrEmpty(value.ProviderReference) &&
                    !Regex.IsMatch(value.ProviderReference,"^Organization/[A-Za-z0-9][A-Za-z0-9.-]{0,63}$")) throw new FormatException();
                if (value.Enabled && (string.IsNullOrWhiteSpace(value.ClientId) || string.IsNullOrWhiteSpace(value.BaseUrl) ||
                    string.IsNullOrWhiteSpace(value.TokenUrl) ||
                    (string.IsNullOrWhiteSpace(value.DocumentTypeCode) && string.IsNullOrWhiteSpace(value.DocumentTypeDisplay)))) throw new FormatException();
                value.AttachmentReadbackRoots=AriaAttachmentReader.ValidateRoots(value.AttachmentReadbackRoots);
                return value;
            }
            catch (Exception error) when (error is JsonException || error is FormatException || error is ArgumentException || error is OverflowException)
            { throw new FormatException("ARIA upload configuration is invalid or unsupported. No secret values are accepted in this JSON."); }
        }

        public static async Task<AriaUploadConfiguration> LoadAsync(string path, CancellationToken cancellation)
        { return Parse(await ReadTextAsync(path,32768,cancellation).ConfigureAwait(false)); }

        public async Task<string> LoadSecretAsync(string configurationPath, CancellationToken cancellation)
        {
            if (!Enabled) throw new InvalidOperationException("ARIA upload is disabled.");
            cancellation.ThrowIfCancellationRequested();
            string secret=Environment.GetEnvironmentVariable(CredentialEnvironmentVariable);
            if (string.IsNullOrEmpty(secret) && !string.IsNullOrWhiteSpace(CredentialEnvFilePath))
            {
                string path=Path.IsPathRooted(CredentialEnvFilePath) ? CredentialEnvFilePath :
                    Path.Combine(Path.GetDirectoryName(Path.GetFullPath(configurationPath)),CredentialEnvFilePath);
                string content=await ReadTextAsync(path,8192,cancellation).ConfigureAwait(false);
                foreach (string raw in content.Split(new[]{'\r','\n'},StringSplitOptions.RemoveEmptyEntries))
                {
                    string line=raw.Trim();
                    if (line.StartsWith("#",StringComparison.Ordinal)) continue;
                    int delimiter=line.IndexOf('=');
                    if (delimiter<=0 || line.Substring(0,delimiter).Trim()!=CredentialEnvironmentVariable) continue;
                    if (secret!=null) throw new InvalidOperationException("Credential file contains a duplicate configured entry.");
                    secret=line.Substring(delimiter+1).Trim();
                    if (secret.Length>=2 && ((secret[0]=='\"' && secret[secret.Length-1]=='\"') ||
                        (secret[0]=='\'' && secret[secret.Length-1]=='\''))) secret=secret.Substring(1,secret.Length-2);
                }
            }
            if (string.IsNullOrWhiteSpace(secret) || secret.Length>4096 || secret.IndexOfAny(new[]{'\r','\n','\0'})>=0)
                throw new InvalidOperationException("The configured ARIA credential is missing or invalid.");
            return secret;
        }

        internal static Uri ValidateHttps(string text, bool baseUrl)
        {
            Uri uri;
            if (text==null || text!=text.Trim() || !Uri.TryCreate(text,UriKind.Absolute,out uri) || uri.Scheme!=Uri.UriSchemeHttps ||
                uri.UserInfo.Length!=0 || uri.Query.Length!=0 || uri.Fragment.Length!=0 || text.IndexOf('\\')>=0 ||
                (baseUrl && uri.AbsolutePath=="/")) throw new FormatException();
            return uri;
        }
        private static bool OptionalText(string text,int maximum)
        { return text==null || (text.Length<=maximum && text==text.Trim() && !Regex.IsMatch(text,"[\\p{Cc}\\p{Zl}\\p{Zp}]")); }
        private static bool ValidPin(string value)
        { return string.IsNullOrEmpty(value) || Regex.IsMatch(value,"^[0-9a-fA-F]{64}$"); }

        private static async Task<string> ReadTextAsync(string path,int maximumBytes,CancellationToken cancellation)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new InvalidOperationException("An external ARIA configuration path is required.");
            cancellation.ThrowIfCancellationRequested();
            Task<string> read=Task.Run(()=> {
                try
                {
                    using(var stream=new FileStream(path,FileMode.Open,FileAccess.Read,FileShare.ReadWrite))
                    {
                        if (stream.Length>maximumBytes) throw new InvalidOperationException("ARIA configuration exceeds the size limit.");
                        var bytes=new byte[maximumBytes+1]; int count=0,received;
                        while ((received=stream.Read(bytes,count,bytes.Length-count))>0)
                        { count+=received; if(count>maximumBytes) throw new InvalidOperationException("ARIA configuration exceeds the size limit."); }
                        int start=count>=3 && bytes[0]==239 && bytes[1]==187 && bytes[2]==191 ? 3 : 0;
                        return new UTF8Encoding(false,true).GetString(bytes,start,count-start);
                    }
                }
                catch (Exception error) when (error is IOException || error is UnauthorizedAccessException || error is ArgumentException || error is System.Security.SecurityException)
                { throw new InvalidOperationException("The external ARIA configuration or credential file cannot be read."); }
            });
            var delay=Task.Delay(TimeSpan.FromSeconds(10),cancellation);
            if (await Task.WhenAny(read,delay).ConfigureAwait(false)!=read)
            {
                _=read.ContinueWith(task=> { var observed=task.Exception; },TaskContinuationOptions.OnlyOnFaulted);
                cancellation.ThrowIfCancellationRequested();
                throw new TimeoutException("ARIA configuration loading timed out.");
            }
            cancellation.ThrowIfCancellationRequested();
            return await read.ConfigureAwait(false);
        }
    }
}
