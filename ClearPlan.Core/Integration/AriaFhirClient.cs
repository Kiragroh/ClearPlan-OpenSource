using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ClearPlan.Core.Integration
{
    public sealed class AriaFhirException : Exception
    {
        internal AriaFhirException(string message) : base(message) { }
    }

    public sealed class AriaResolvedPatient
    {
        public string Identifier { get; internal set; }
        public string Reference { get; internal set; }
    }

    public sealed class AriaResolvedProvider
    {
        public string Reference { get; internal set; }
        public string Display { get; internal set; }
    }

    public sealed class AriaResolvedDocumentType
    {
        public string System { get; internal set; }
        public string Code { get; internal set; }
        public string Display { get; internal set; }
    }

    public enum AriaUploadState { Sent, Uncertain, Rejected }
    public sealed class AriaUploadResult
    {
        public AriaUploadState State { get; internal set; }
        public string ResourceReference { get; internal set; }
        public string Message { get; internal set; }
        public int? HttpStatusCode { get; internal set; }
        public string BusinessRuleCode { get; internal set; }
        public string ProcessingStageCode { get; internal set; }
    }

    public enum AriaVerificationState { Verified, MetadataOnly, Mismatch, Unavailable }
    public sealed class AriaVerificationResult
    {
        public AriaVerificationState State { get; internal set; }
        public string Message { get; internal set; }
        public bool AttachmentCreationChecked { get; internal set; }
    }

    /// <summary>Uses a caller-owned HttpClient with redirects disabled; never retries a write.</summary>
    public sealed class AriaFhirClient
    {
        private const string TypeValueSet = "http://varian.com/fhir/ValueSet/documentreference-type";
        private const string CategorySystem = "http://varian.com/fhir/CodeSystem/DocumentReference/documentreference-class";
        private const int MaximumLookupBytes = 1024 * 1024;
        private const int MaximumDocumentBytes = 29 * 1024 * 1024;
        private const int MaximumPages = 20;
        private readonly HttpClient http;
        private readonly Uri serviceBase;
        private readonly Func<CancellationToken, Task<string>> tokenProvider;
        private readonly int timeoutSeconds;
        private readonly Func<string, CancellationToken, Task<byte[]>> attachmentReader;

        public AriaFhirClient(HttpClient http, Uri baseUri, Func<CancellationToken, Task<string>> tokenProvider,
            int requestTimeoutSeconds = 45)
            : this(http, baseUri, tokenProvider, requestTimeoutSeconds, null) { }

        public AriaFhirClient(HttpClient http, Uri baseUri, Func<CancellationToken, Task<string>> tokenProvider,
            int requestTimeoutSeconds, Func<string, CancellationToken, Task<byte[]>> attachmentReader = null)
        {
            if (http == null || tokenProvider == null || baseUri == null || !baseUri.IsAbsoluteUri ||
                baseUri.Scheme != Uri.UriSchemeHttps || baseUri.UserInfo.Length != 0 || baseUri.Query.Length != 0 ||
                baseUri.Fragment.Length != 0 || baseUri.AbsolutePath == "/" || requestTimeoutSeconds < 5 || requestTimeoutSeconds > 120)
                throw Error("The FHIR connection configuration is invalid.");
            this.http = http;
            this.tokenProvider = tokenProvider;
            timeoutSeconds = requestTimeoutSeconds;
            this.attachmentReader = attachmentReader;
            serviceBase = new Uri(baseUri.AbsoluteUri.TrimEnd('/') + "/");
            RequireTarget(serviceBase.AbsoluteUri);
        }

        public Task<AriaResolvedPatient> ResolvePatientAsync(string identifier, string identifierSystem, CancellationToken cancellation)
        {
            return ReadOperation(async token => {
                RequireText(identifier, 256);
                if (identifier.IndexOfAny(new[] { '|', ',', '\\', '$' }) >= 0) throw Error("The patient identifier contains unsupported search characters.");
                if (!string.IsNullOrEmpty(identifierSystem))
                {
                    RequireText(identifierSystem, 512);
                    if (identifierSystem.IndexOfAny(new[] { '|', ',', '\\', '$' }) >= 0) throw Error("The patient identifier system is invalid.");
                }
                string search = string.IsNullOrEmpty(identifierSystem) ? identifier : identifierSystem + "|" + identifier;
                var resources = await ReadResources("Patient?identifier=" + Uri.EscapeDataString(search), "Patient", token).ConfigureAwait(false);
                var matches = resources.Where(p => p["active"] == null || (bool?)p["active"] != false).Where(p =>
                    Array(p["identifier"]).OfType<JObject>().Any(i => (string)i["value"] == identifier &&
                        (string.IsNullOrEmpty(identifierSystem) || (string)i["system"] == identifierSystem))).ToList();
                if (matches.Count != 1) throw Error("ARIA patient lookup did not return exactly one matching patient.");
                return new AriaResolvedPatient { Identifier = identifier, Reference = "Patient/" + ResourceId(matches[0]) };
            }, cancellation);
        }

        public Task<AriaResolvedProvider> ResolveProviderAsync(string configuredReference, CancellationToken cancellation)
        {
            return ReadOperation(async token => {
                List<JObject> resources;
                if (!string.IsNullOrEmpty(configuredReference))
                {
                    RequireRelativeReference(configuredReference, "Organization");
                    var organization = await GetJson(RequireTarget(configuredReference), MaximumLookupBytes, token).ConfigureAwait(false);
                    if ((string)organization["resourceType"] != "Organization" || "Organization/" + ResourceId(organization) != configuredReference)
                        throw Error("The configured provider reference could not be verified.");
                    resources = new List<JObject> { organization };
                }
                else resources = await ReadResources("Organization?type=prov&active=true", "Organization", token).ConfigureAwait(false);
                var matches = resources.Where(o => (bool?)o["active"] == true && Array(o["type"]).OfType<JObject>().Any(t =>
                    Array(t["coding"]).OfType<JObject>().Any(c => (string)c["code"] == "prov" &&
                        (string)c["system"] == "http://terminology.hl7.org/CodeSystem/organization-type"))).ToList();
                if (matches.Count != 1) throw Error("ARIA provider lookup did not return exactly one active provider.");
                string display = (string)matches[0]["name"];
                RequireText(display, 256);
                return new AriaResolvedProvider { Reference = "Organization/" + ResourceId(matches[0]), Display = display };
            }, cancellation);
        }

        public Task<AriaResolvedDocumentType> ResolveDocumentTypeAsync(string providerReference, string configuredCode,
            string configuredDisplay, CancellationToken cancellation)
        {
            return ReadOperation(async token => {
                RequireRelativeReference(providerReference, "Organization");
                if (string.IsNullOrEmpty(configuredCode) && string.IsNullOrEmpty(configuredDisplay))
                    throw Error("An exact document type code or display must be configured.");
                if (!string.IsNullOrEmpty(configuredCode)) RequireText(configuredCode, 256);
                if (!string.IsNullOrEmpty(configuredDisplay)) RequireText(configuredDisplay, 256);
                var values = await ReadResources("ValueSet/$expand?url=" + Uri.EscapeDataString(TypeValueSet) +
                    "&publisher=" + Uri.EscapeDataString(providerReference.Substring("Organization/".Length)), "ValueSet", token).ConfigureAwait(false);
                var codes = new List<AriaResolvedDocumentType>();
                int count = 0;
                int declaredTotal = 0;
                foreach (var value in values)
                {
                    if ((string)value["url"] != TypeValueSet || !(value["expansion"] is JObject))
                        throw Error("ARIA did not return the requested expanded document terminology.");
                    var expansion = (JObject)value["expansion"];
                    int? total = (int?)expansion["total"];
                    int? offset = (int?)expansion["offset"];
                    if (total.HasValue && total.Value < 0) throw Error("The document terminology total is invalid.");
                    if (offset.HasValue && offset.Value != count) throw Error("The document terminology expansion skips or repeats entries.");
                    declaredTotal = Math.Max(declaredTotal, total ?? 0);
                    CollectCodes(Array(expansion["contains"]), null, false, codes, ref count);
                }
                if (declaredTotal > count) throw Error("The document terminology expansion is incomplete.");
                var uniqueCodes = codes.GroupBy(c => c.System + "\n" + c.Code, StringComparer.Ordinal).Select(g => {
                        if (g.Select(c => c.Display).Distinct(StringComparer.Ordinal).Count() != 1)
                            throw Error("The document terminology contains conflicting entries.");
                        return g.First();
                    }).ToList();
                var matches = uniqueCodes.Where(c => (string.IsNullOrEmpty(configuredCode) || c.Code == configuredCode) &&
                    (string.IsNullOrEmpty(configuredDisplay) || c.Display == configuredDisplay)).ToList();
                if (matches.Count != 1) throw Error("ARIA document type lookup did not return exactly one exact match.");
                return matches[0];
            }, cancellation);
        }

        public async Task<AriaUploadResult> UploadAsync(AriaReportUploadRequest request, CancellationToken cancellation)
        {
            string stage = "preparing";
            bool submitted = false;
            int? receivedStatus = null;
            string recoveryReference = null;
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellation))
            {
                timeout.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds));
                try
                {
                    timeout.Token.ThrowIfCancellationRequested();
                    JObject payload = AriaReportDocumentBuilder.Build(request);
                    Uri target = RequireTarget("DocumentReference");
                    using (var message = new HttpRequestMessage(HttpMethod.Post, target))
                    {
                        message.Content = new StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/fhir+json");
                        message.Headers.TryAddWithoutValidation("Prefer", "return=representation");
                        await Authorize(message, timeout.Token).ConfigureAwait(false);
                        timeout.Token.ThrowIfCancellationRequested();
                        // From this point onward, a transport failure cannot establish whether ARIA stored the document.
                        submitted = true;
                        stage = "sending";
                        using (var response = await AwaitBounded(http.SendAsync(message, HttpCompletionOption.ResponseHeadersRead, timeout.Token), timeout.Token).ConfigureAwait(false))
                        {
                            int status = (int)response.StatusCode;
                            receivedStatus = status;
                            if (status >= 400 && status < 500 && status != 408)
                            {
                                string businessRule = await ReadSafeBusinessRule(response.Content, timeout.Token).ConfigureAwait(false);
                                string rejection = businessRule == "FUTURE_DATE_TIME"
                                    ? "ARIA rejected the document (FUTURE_DATE_TIME). Check the ARIA document date, workstation/server clocks and DocumentDateSafetyMinutes. Report creation is unchanged; no automatic retry was performed."
                                    : "ARIA rejected the document. No automatic retry was performed.";
                                return UploadResult(AriaUploadState.Rejected, null, rejection + " [HTTP " + status + "]", status, businessRule, "server-rejection");
                            }
                            if (status != 200 && status != 201)
                                return UploadResult(AriaUploadState.Uncertain, null, "ARIA storage is unconfirmed. Reconcile before sending again.", status, processingStageCode: "response-status");
                            string locationReference = null;
                            stage = "create-location";
                            if (response.Headers.Location != null)
                            {
                                locationReference = DocumentReference(RequireTarget(response.Headers.Location.OriginalString), true);
                                recoveryReference = locationReference;
                            }
                            // A trusted Location is only a recovery candidate. A response-target failure can never become Sent.
                            stage = "response-target";
                            bool responseTargetValid = true;
                            try { RequireResponseTarget(response, target); }
                            catch (AriaFhirException) { responseTargetValid = false; }
                            // Keep the bounded body consistency check even after a target failure; conflicting IDs discard the candidate.
                            stage = "create-body";
                            JObject body = await ReadJson(response.Content, MaximumDocumentBytes, timeout.Token, true).ConfigureAwait(false);
                            string bodyReference = null;
                            if (body != null)
                            {
                                if ((string)body["resourceType"] != "DocumentReference") throw Error("ARIA returned an unexpected create response.");
                                bodyReference = "DocumentReference/" + ResourceId(body);
                            }
                            stage = "create-reference";
                            if (bodyReference != null && locationReference != null && bodyReference != locationReference)
                            { recoveryReference = null; throw Error("ARIA create response references disagree."); }
                            if (!responseTargetValid)
                                return UploadResult(AriaUploadState.Uncertain, recoveryReference,
                                    "ARIA response target was unexpected. Storage is unconfirmed; reconcile before sending again.",
                                    status, processingStageCode: "response-target");
                            string reference = bodyReference ?? locationReference;
                            if (reference == null) throw Error("ARIA returned no verifiable document reference.");
                            return UploadResult(AriaUploadState.Sent, reference, "ARIA accepted the document; readback verification is required.", status, processingStageCode: "accepted");
                        }
                    }
                }
                catch (Exception)
                {
                    return UploadResult(submitted ? AriaUploadState.Uncertain : AriaUploadState.Rejected, recoveryReference,
                        submitted ? "ARIA storage is unconfirmed. Reconcile before sending again." : "The document was not submitted. Check configuration and authorization.", receivedStatus, processingStageCode: stage);
                }
            }
        }

        public Task<AriaVerificationResult> VerifyAsync(string resourceReference, AriaReportUploadRequest request, CancellationToken cancellation)
        {
            return VerifyCore(resourceReference, request, cancellation);
        }

        private async Task<AriaVerificationResult> VerifyCore(string reference, AriaReportUploadRequest request, CancellationToken cancellation)
        {
            try
            {
                return await ReadOperation(async token => {
                    if (request == null) throw Error("A prepared document is required for verification.");
                    RequireRelativeReference(reference, "DocumentReference");
                    var document = await GetJson(RequireTarget(reference), MaximumDocumentBytes, token).ConfigureAwait(false);
                    bool metadata = (string)document["resourceType"] == "DocumentReference" && "DocumentReference/" + ResourceId(document) == reference &&
                        (string)document["status"] == "current" && (string)document["docStatus"] == "preliminary" &&
                        (string)document["subject"]?["reference"] == request.PatientReference &&
                        (string)document["custodian"]?["reference"] == request.OrganizationReference &&
                        HasCoding(document["type"], request.DocumentTypeSystem, request.DocumentTypeCode) &&
                        Array(document["category"]).Any(c => HasCoding(c, CategorySystem, request.CategoryCode));
                    var contents = Array(document["content"]);
                    var attachment = contents.Count == 1 ? contents[0]["attachment"] as JObject : null;
                    bool titlePresent = attachment != null && attachment["title"] != null && attachment["title"].Type != JTokenType.Null;
                    bool titleOrDescription = titlePresent ? (string)attachment["title"] == request.Title :
                        !string.IsNullOrEmpty(request.Description) && (string)document["description"] == request.Description;
                    if (!metadata || attachment == null || (string)attachment["contentType"] != "application/pdf" || !titleOrDescription ||
                        !MatchesInstant(document["date"], request.DocumentDateUtc))
                        return VerificationResult(AriaVerificationState.Mismatch, "ARIA readback does not match the prepared report metadata.");
                    bool creationPresent = attachment["creation"] != null && attachment["creation"].Type != JTokenType.Null;
                    if (creationPresent && !MatchesInstant(attachment["creation"], request.CreatedUtc))
                        return VerificationResult(AriaVerificationState.Mismatch, "ARIA attachment creation differs from the prepared report.");
                    if (attachment["size"] != null && (long)attachment["size"] != request.PdfLength)
                        return VerificationResult(AriaVerificationState.Mismatch, "ARIA readback document size differs.");
                    bool verified = false;
                    string data = (string)attachment["data"];
                    if (!string.IsNullOrEmpty(data))
                    {
                        byte[] bytes = Convert.FromBase64String(data);
                        if (bytes.Length != request.PdfLength || Sha256(bytes) != request.PdfSha256)
                            return VerificationResult(AriaVerificationState.Mismatch, "ARIA readback document bytes differ.");
                        verified = true;
                    }
                    string hash = (string)attachment["hash"];
                    if (!string.IsNullOrEmpty(hash))
                    {
                        byte[] expected;
                        using (var sha = SHA1.Create()) expected = sha.ComputeHash(request.PdfBytes);
                        if (!expected.SequenceEqual(Convert.FromBase64String(hash)))
                            return VerificationResult(AriaVerificationState.Mismatch, "ARIA readback document fingerprint differs.");
                        verified = true;
                    }
                    string attachmentUrl = (string)attachment["url"];
                    if (!verified && attachmentReader != null && attachmentUrl != null && attachmentUrl.StartsWith("\\\\", StringComparison.Ordinal))
                    {
                        byte[] bytes = await AwaitBounded(attachmentReader(attachmentUrl, token), token).ConfigureAwait(false);
                        if (bytes == null || bytes.Length != request.PdfLength || Sha256(bytes) != request.PdfSha256)
                            return VerificationResult(AriaVerificationState.Mismatch, "Stored attachment bytes differ from the prepared PDF.");
                        verified = true;
                    }
                    var result = VerificationResult(verified ? AriaVerificationState.Verified : AriaVerificationState.MetadataOnly,
                        (verified ? "ARIA report identity, document date, preliminary status and PDF fingerprint verified." :
                        "ARIA report metadata verified; attachment bytes or a fingerprint could not be verified.") +
                        (creationPresent ? " Attachment creation verified." : " Attachment creation was not returned by ARIA and was not checked."));
                    result.AttachmentCreationChecked = creationPresent;
                    return result;
                }, cancellation).ConfigureAwait(false);
            }
            catch (Exception)
            { return VerificationResult(AriaVerificationState.Unavailable, "ARIA readback could not be verified. Do not resend automatically."); }
        }

        private async Task<T> ReadOperation<T>(Func<CancellationToken, Task<T>> operation, CancellationToken cancellation)
        {
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellation))
            {
                timeout.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds));
                try { timeout.Token.ThrowIfCancellationRequested(); return await AwaitBounded(operation(timeout.Token), timeout.Token).ConfigureAwait(false); }
                catch (OperationCanceledException)
                {
                    if (cancellation.IsCancellationRequested) throw new OperationCanceledException(cancellation);
                    throw Error("ARIA read timed out; no document was sent by this operation.");
                }
                catch (AriaFhirException) { throw; }
                catch (Exception) { throw Error("ARIA returned an invalid or unavailable response. No server response text is included."); }
            }
        }

        private async Task<List<JObject>> ReadResources(string path, string resourceType, CancellationToken cancellation)
        {
            Uri next = RequireTarget(path);
            var visited = new HashSet<string>(StringComparer.Ordinal);
            var resources = new List<JObject>();
            var identifiers = new Dictionary<string, JObject>(StringComparer.Ordinal);
            for (int page = 0; next != null; page++)
            {
                if (page >= MaximumPages || !visited.Add(next.AbsoluteUri)) throw Error("ARIA pagination exceeded its safe limit or repeated a page.");
                var document = await GetJson(next, MaximumLookupBytes, cancellation).ConfigureAwait(false);
                string type = (string)document["resourceType"];
                if (type == "ValueSet" && resourceType == "ValueSet")
                { resources.Add(document); next = null; continue; }
                if (type != "Bundle") throw Error("ARIA lookup did not return a resource bundle.");
                foreach (var entry in Array(document["entry"]).OfType<JObject>())
                {
                    var resource = entry["resource"] as JObject;
                    if (resource == null) throw Error("ARIA lookup contains a missing resource.");
                    if ((string)resource["resourceType"] == "OperationOutcome") throw Error("ARIA returned a lookup issue; details are not included.");
                    if ((string)entry["search"]?["mode"] == "include") continue;
                    if ((string)resource["resourceType"] != resourceType) throw Error("ARIA lookup returned an unexpected resource type.");
                    if (resourceType == "ValueSet") resources.Add(resource);
                    else
                    {
                        string id = ResourceId(resource);
                        JObject previous;
                        if (identifiers.TryGetValue(id, out previous))
                        {
                            // The same ID on two pages is safe only if identity/provider metadata is unchanged.
                            if (!JToken.DeepEquals(previous["identifier"], resource["identifier"]) ||
                                !JToken.DeepEquals(previous["active"], resource["active"]) || !JToken.DeepEquals(previous["type"], resource["type"]))
                                throw Error("ARIA lookup returned conflicting versions of a resource.");
                        }
                        else { identifiers.Add(id, resource); resources.Add(resource); }
                    }
                    if (resources.Count > 10000) throw Error("ARIA lookup returned too many resources.");
                }
                var links = Array(document["link"]).OfType<JObject>().Where(l => (string)l["relation"] == "next").ToList();
                if (links.Count > 1) throw Error("ARIA returned ambiguous pagination links.");
                next = links.Count == 0 ? null : RequireTarget((string)links[0]["url"]);
                if (next == null && resourceType != "ValueSet" && document["total"] != null && (int)document["total"] > resources.Count)
                    throw Error("ARIA lookup pagination is incomplete.");
            }
            return resources;
        }

        private async Task<JObject> GetJson(Uri target, int maximumBytes, CancellationToken cancellation)
        {
            using (var message = new HttpRequestMessage(HttpMethod.Get, target))
            {
                await Authorize(message, cancellation).ConfigureAwait(false);
                using (var response = await AwaitBounded(http.SendAsync(message, HttpCompletionOption.ResponseHeadersRead, cancellation), cancellation).ConfigureAwait(false))
                {
                    RequireResponseTarget(response, target);
                    if (!response.IsSuccessStatusCode) throw Error("ARIA rejected the read request. Check authorization and configuration.");
                    return await ReadJson(response.Content, maximumBytes, cancellation, false).ConfigureAwait(false);
                }
            }
        }

        private async Task Authorize(HttpRequestMessage request, CancellationToken cancellation)
        {
            RequireTarget(request.RequestUri.AbsoluteUri);
            cancellation.ThrowIfCancellationRequested();
            string token = await AwaitBounded(tokenProvider(cancellation), cancellation).ConfigureAwait(false);
            cancellation.ThrowIfCancellationRequested();
            if (string.IsNullOrEmpty(token) || token.Length > 32768 || !Regex.IsMatch(token, @"\A[A-Za-z0-9\-._~+/]+=*\z"))
                throw Error("ARIA authorization did not return a valid bearer token.");
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/fhir+json"));
        }

        private static async Task<string> ReadSafeBusinessRule(HttpContent content, CancellationToken cancellation)
        {
            // The HTTP rejection is already known. A malformed, large or stalled error body must not change it to Uncertain.
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellation))
            {
                timeout.CancelAfter(TimeSpan.FromSeconds(2));
                try
                {
                    var outcome = await ReadJson(content, 65536, timeout.Token, true, true).ConfigureAwait(false);
                    if (outcome == null || (string)outcome["resourceType"] != "OperationOutcome") return null;
                    foreach (var issue in Array(outcome["issue"]).OfType<JObject>().Take(32))
                    {
                        var details = issue["details"] as JObject;
                        var values = new List<string> { (string)issue["diagnostics"] };
                        if (details != null) values.AddRange(Array(details["coding"]).OfType<JObject>().Select(c => (string)c["code"]));
                        if (values.Any(value => value != null && Regex.IsMatch(value, @"(?<![A-Za-z0-9_])FUTURE_DATE_TIME(?![A-Za-z0-9_])")))
                            return "FUTURE_DATE_TIME";
                    }
                }
                catch (Exception) { /* Only a literal allowlisted code may escape the bounded error parser. */ }
                return null;
            }
        }

        private static async Task<JObject> ReadJson(HttpContent content, int maximumBytes, CancellationToken cancellation, bool allowEmpty,
            bool allowOperationOutcome = false)
        {
            if (content == null) { if (allowEmpty) return null; throw Error("ARIA returned an empty response."); }
            if (content.Headers.ContentLength > maximumBytes) throw Error("ARIA response exceeds the size limit.");
            using (var stream = await AwaitBounded(content.ReadAsStreamAsync(), cancellation).ConfigureAwait(false))
            using (var memory = new MemoryStream())
            {
                var buffer = new byte[8192]; int count;
                while ((count = await AwaitBounded(stream.ReadAsync(buffer, 0, buffer.Length, cancellation), cancellation).ConfigureAwait(false)) > 0)
                {
                    if (memory.Length + count > maximumBytes) throw Error("ARIA response exceeds the size limit.");
                    memory.Write(buffer, 0, count);
                }
                cancellation.ThrowIfCancellationRequested();
                string json = new UTF8Encoding(false, true).GetString(memory.ToArray());
                if (string.IsNullOrWhiteSpace(json)) { if (allowEmpty) return null; throw Error("ARIA returned an empty response."); }
                using (var reader = new JsonTextReader(new StringReader(json)) { MaxDepth = 32, DateParseHandling = DateParseHandling.None })
                {
                    var value = JObject.Load(reader, new JsonLoadSettings { DuplicatePropertyNameHandling = DuplicatePropertyNameHandling.Error });
                    if (reader.Read() || value.Descendants().Take(50001).Count() > 50000) throw Error("ARIA response JSON exceeds its safe limit.");
                    if (!allowOperationOutcome && (string)value["resourceType"] == "OperationOutcome") throw Error("ARIA returned an operation issue; details are not included.");
                    return value;
                }
            }
        }

        private Uri RequireTarget(string reference)
        {
            Uri value;
            if (string.IsNullOrWhiteSpace(reference) || reference != reference.Trim() || reference.Length > 4096 ||
                reference.Any(char.IsControl) || reference.IndexOf('\\') >= 0 || !Uri.TryCreate(serviceBase, reference, out value) ||
                value.Scheme != Uri.UriSchemeHttps || value.Host != serviceBase.Host || value.Port != serviceBase.Port ||
                value.UserInfo.Length != 0 || value.Fragment.Length != 0 || !value.AbsolutePath.StartsWith(serviceBase.AbsolutePath, StringComparison.Ordinal) ||
                Regex.IsMatch(value.GetComponents(UriComponents.Path, UriFormat.UriEscaped), "%2e|%2f|%5c|%25", RegexOptions.IgnoreCase))
                throw Error("ARIA link is outside the configured HTTPS service boundary.");
            return value;
        }

        private static async Task<T> AwaitBounded<T>(Task<T> task, CancellationToken cancellation)
        {
            var cancelled = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            using (cancellation.Register(() => cancelled.TrySetResult(true)))
            {
                if (await Task.WhenAny(task, cancelled.Task).ConfigureAwait(false) == task)
                    return await task.ConfigureAwait(false);
                // A supplied dependency may ignore cancellation. Observe eventual failures and dispose any late stream/response.
                _ = task.ContinueWith(completed => {
                    if (completed.IsFaulted) { var observed = completed.Exception; }
                    else if (completed.Status == TaskStatus.RanToCompletion) (completed.Result as IDisposable)?.Dispose();
                }, CancellationToken.None, TaskContinuationOptions.ExecuteSynchronously, TaskScheduler.Default);
                throw new OperationCanceledException(cancellation);
            }
        }

        private void RequireResponseTarget(HttpResponseMessage response, Uri expected)
        {
            if (response.RequestMessage != null && response.RequestMessage.RequestUri != null &&
                RequireTarget(response.RequestMessage.RequestUri.AbsoluteUri) != expected)
                throw Error("ARIA redirected a request unexpectedly.");
        }

        private string DocumentReference(Uri uri, bool allowHistory)
        {
            string relative = uri.AbsolutePath.Substring(serviceBase.AbsolutePath.Length);
            // Extract only a strict canonical resource identity; a Location query is never retained or followed.
            if (!Regex.IsMatch(relative, @"\ADocumentReference/[A-Za-z0-9.-]{1,64}" +
                (allowHistory ? @"(?:/_history/[A-Za-z0-9.-]{1,64})?\z" : @"\z")))
                throw Error("ARIA returned an invalid document reference.");
            string reference = string.Join("/", relative.Split('/').Take(2));
            RequireRelativeReference(reference, "DocumentReference");
            return reference;
        }

        private static void CollectCodes(JArray entries, string inheritedSystem, bool inheritedInactive,
            List<AriaResolvedDocumentType> codes, ref int count)
        {
            foreach (var entry in entries.OfType<JObject>())
            {
                string system = (string)entry["system"] ?? inheritedSystem;
                bool inactive = inheritedInactive || (bool?)entry["inactive"] == true;
                string code = (string)entry["code"];
                if (code != null)
                {
                    if (++count > 10000) throw Error("ARIA terminology contains too many codes.");
                    if (!inactive && (bool?)entry["abstract"] != true)
                    {
                        RequireText(system, 512); RequireText(code, 256);
                        string display = (string)entry["display"]; RequireText(display, 256);
                        Uri systemUri;
                        if (!Uri.TryCreate(system, UriKind.Absolute, out systemUri) ||
                            (systemUri.Scheme != "http" && systemUri.Scheme != "https" && systemUri.Scheme != "urn") ||
                            systemUri.UserInfo.Length != 0 || systemUri.Query.Length != 0 || systemUri.Fragment.Length != 0)
                            throw Error("ARIA terminology contains an invalid coding system.");
                        codes.Add(new AriaResolvedDocumentType { System = system, Code = code, Display = display });
                    }
                }
                CollectCodes(Array(entry["contains"]), system, inactive, codes, ref count);
            }
        }

        private static bool HasCoding(JToken concept, string system, string code)
        { return concept is JObject && Array(concept["coding"]).OfType<JObject>().Any(c => (string)c["system"] == system && (string)c["code"] == code); }
        private static bool MatchesInstant(JToken value, DateTimeOffset expected)
        {
            string text = value == null ? null : (string)value;
            DateTimeOffset instant;
            return text != null && Regex.IsMatch(text, @"\A\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})\z") &&
                DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out instant) &&
                Math.Abs((instant - expected).Ticks) <= TimeSpan.FromMilliseconds(3).Ticks;
        }
        private static JArray Array(JToken token)
        { if (token == null || token.Type == JTokenType.Null) return new JArray(); var array = token as JArray; if (array == null) throw Error("ARIA returned an invalid resource collection."); return array; }
        private static string ResourceId(JObject resource)
        { string id = (string)resource["id"]; if (id == "." || id == ".." || id == null || !Regex.IsMatch(id, @"\A[A-Za-z0-9.-]{1,64}\z")) throw Error("ARIA returned an invalid resource identifier."); return id; }
        private static void RequireRelativeReference(string value, string type)
        { if (value == null || !Regex.IsMatch(value, "\\A" + type + @"/[A-Za-z0-9.-]{1,64}\z") || value == type + "/." || value == type + "/..") throw Error("A resolved relative FHIR reference is required."); }
        private static void RequireText(string value, int maximum)
        { if (string.IsNullOrWhiteSpace(value) || value.Length > maximum || value != value.Trim() || Regex.IsMatch(value, @"[\p{Cc}\p{Cf}\p{Zl}\p{Zp}<>]")) throw Error("ARIA lookup metadata is missing or invalid."); }
        private static string Sha256(byte[] bytes)
        { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant(); }
        private static AriaFhirException Error(string message) { return new AriaFhirException(message); }
        private static AriaUploadResult UploadResult(AriaUploadState state, string reference, string message,
            int? httpStatusCode = null, string businessRuleCode = null, string processingStageCode = null)
        { return new AriaUploadResult { State = state, ResourceReference = reference, Message = message,
            HttpStatusCode = httpStatusCode, BusinessRuleCode = businessRuleCode, ProcessingStageCode = processingStageCode }; }
        private static AriaVerificationResult VerificationResult(AriaVerificationState state, string message)
        { return new AriaVerificationResult { State = state, Message = message }; }
    }
}
