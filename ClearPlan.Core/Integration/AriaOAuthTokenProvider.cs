using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ClearPlan.Core.Integration
{
    /// <summary>Client credentials stay in memory. The supplied HTTP client must reject redirects.</summary>
    public sealed class AriaOAuthTokenProvider
    {
        private readonly HttpClient http;
        private readonly Uri endpoint;
        private readonly string clientId, scope;
        private readonly Func<CancellationToken,Task<string>> secretProvider;
        private readonly SemaphoreSlim gate=new SemaphoreSlim(1,1);
        private string token;
        private DateTime expires=DateTime.MinValue;

        public AriaOAuthTokenProvider(HttpClient http,Uri endpoint,string clientId,string scope,
            Func<CancellationToken,Task<string>> secretProvider)
        {
            this.http=http ?? throw new ArgumentNullException("http");
            this.endpoint=AriaUploadConfiguration.ValidateHttps(endpoint==null ? null : endpoint.AbsoluteUri,false);
            if(string.IsNullOrWhiteSpace(clientId) || string.IsNullOrWhiteSpace(scope) || secretProvider==null)
                throw new ArgumentException("Explicit OAuth configuration and an external credential provider are required.");
            this.clientId=clientId; this.scope=scope; this.secretProvider=secretProvider;
        }

        public async Task<string> GetTokenAsync(CancellationToken cancellation)
        {
            await gate.WaitAsync(cancellation).ConfigureAwait(false);
            try
            {
                cancellation.ThrowIfCancellationRequested();
                if(token!=null && DateTime.UtcNow<expires) return token;
                using(var limit=CancellationTokenSource.CreateLinkedTokenSource(cancellation))
                {
                    limit.CancelAfter(TimeSpan.FromSeconds(60));
                    string secret=await secretProvider(limit.Token).ConfigureAwait(false);
                    using(var request=new HttpRequestMessage(HttpMethod.Post,endpoint))
                    {
                        request.Content=new FormUrlEncodedContent(new Dictionary<string,string> {
                            {"grant_type","client_credentials"},{"client_id",clientId},{"client_secret",secret},{"scope",scope} });
                        using(var response=await http.SendAsync(request,HttpCompletionOption.ResponseHeadersRead,limit.Token).ConfigureAwait(false))
                        {
                            if(!response.IsSuccessStatusCode || response.RequestMessage!=null && response.RequestMessage.RequestUri!=endpoint)
                                throw new InvalidOperationException();
                            if(response.Content.Headers.ContentLength>65536) throw new InvalidOperationException();
                            using(var stream=await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                            using(var buffer=new MemoryStream())
                            {
                                byte[] chunk=new byte[4096]; int read;
                                while((read=await stream.ReadAsync(chunk,0,chunk.Length,limit.Token).ConfigureAwait(false))!=0)
                                {
                                    if(buffer.Length+read>65536) throw new InvalidOperationException();
                                    buffer.Write(chunk,0,read);
                                }
                                using(var reader=new JsonTextReader(new StringReader(new UTF8Encoding(false,true).GetString(buffer.ToArray()))){MaxDepth=8})
                                {
                                    var json=JObject.Load(reader);
                                    if(reader.Read()) throw new InvalidOperationException();
                                    string next=(string)json["access_token"],type=(string)json["token_type"];
                                    if(string.IsNullOrWhiteSpace(next) || next.Length>32768 ||
                                        System.Text.RegularExpressions.Regex.IsMatch(next,"\\s|[\\p{Cc}\\p{Zl}\\p{Zp}]") ||
                                        type!=null && !string.Equals(type,"Bearer",StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException();
                                    double duration; string durationText=(string)json["expires_in"];
                                    if(!double.TryParse(durationText,System.Globalization.NumberStyles.Float,System.Globalization.CultureInfo.InvariantCulture,out duration) ||
                                        double.IsNaN(duration) || double.IsInfinity(duration) || duration<0) duration=0;
                                    token=next; expires=DateTime.UtcNow.AddSeconds(Math.Max(0,Math.Min(3600,duration)-30));
                                    return token;
                                }
                            }
                        }
                    }
                }
            }
            catch(OperationCanceledException)
            {
                if(cancellation.IsCancellationRequested) throw new OperationCanceledException(cancellation);
                throw new InvalidOperationException("ARIA authentication timed out. No document was sent.");
            }
            catch(Exception error) when(error is HttpRequestException || error is IOException || error is JsonException ||
                error is InvalidOperationException || error is ArgumentException || error is FormatException)
            { throw new InvalidOperationException("ARIA authentication failed. Check the local connection and credential configuration. No document was sent."); }
            finally { gate.Release(); }
        }
    }
}
