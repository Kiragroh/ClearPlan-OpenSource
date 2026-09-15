using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ClearPlan.Core.Integration;
using Newtonsoft.Json.Linq;

namespace ClearPlan.Core.Tests
{
    internal static class AriaFhirClientTests
    {
        public static void InteractiveAuthorExactLookup()
        {
            const string usernameSystem = "http://varian.com/fhir/identifier/Practitioner/UserName";
            const string idSystem = "http://varian.com/fhir/identifier/Practitioner/Id";
            foreach (string id in new[] { "DOMAIN\\physicist", "physicist" })
            {
                string system = id.Contains("\\") ? usernameSystem : idSystem;
                using (var harness = new Harness(Bundle(Practitioner("a1", id, system))))
                {
                    object author = ResolveAuthor(harness.Client, id);
                    TestAssert.Equal("Practitioner/a1", (string)author.GetType().GetProperty("Reference").GetValue(author));
                    TestAssert.Equal("?identifier=" + Uri.EscapeDataString(system + "|" + id.Replace("\\", "\\\\")), harness.Handler.Uris[0].Query);
                }
            }
            foreach (var response in new[] { Bundle(), Bundle(Practitioner("a1", "OTHER\\physicist", usernameSystem)),
                Bundle(Practitioner("a1", "DOMAIN\\physicist", "urn:wrong")),
                Bundle(Practitioner("a1", "DOMAIN\\physicist", usernameSystem, false)),
                Bundle(Practitioner("a1", "DOMAIN\\physicist", usernameSystem), Practitioner("a2", "DOMAIN\\physicist", usernameSystem)) })
                using (var harness = new Harness(response))
                {
                    var error = TestAssert.Throws<Exception>(() => ResolveAuthor(harness.Client, "DOMAIN\\physicist"));
                    TestAssert.True(error.GetType().Name == "AriaAuthorResolutionException", "Missing/ambiguous identities require an actionable author-specific failure.");
                    TestAssert.False(error.Message.Contains("physicist"));
                    TestAssert.Equal(1, harness.Handler.Uris.Count, "No fuzzy name or service-account fallback is allowed.");
                }
            using (var harness = new Harness())
                foreach (string invalid in new[] { null, "", " physicist", "domain\\", "domain\\user\\other", "user|other", "user,other", "user\n" })
                {
                    TestAssert.Throws<Exception>(() => ResolveAuthor(harness.Client, invalid));
                    TestAssert.Equal(0, harness.Handler.Uris.Count);
                }
            var first = Bundle(Practitioner("a1", "domain\\PHYSICIST", usernameSystem));
            first["link"] = new JArray(new JObject { ["relation"] = "next", ["url"] = "Practitioner?cursor=2" });
            using (var harness = new Harness(first, Bundle()))
            {
                ResolveAuthor(harness.Client, "DOMAIN\\physicist");
                TestAssert.Equal(2, harness.Handler.Uris.Count);
            }
            using (var harness = new Harness(first, Bundle(Practitioner("a2", "DOMAIN\\physicist", usernameSystem))))
                TestAssert.Throws<Exception>(() => ResolveAuthor(harness.Client, "DOMAIN\\physicist"));
            using (var harness = new Harness(first, Bundle(Practitioner("a1", "OTHER\\physicist", usernameSystem))))
                TestAssert.Throws<Exception>(() => ResolveAuthor(harness.Client, "DOMAIN\\physicist"));
        }

        public static void ClinicalMetadataReadbackCannotAcceptServiceAuthor()
        {
            var request = AriaReportUploadTests.BindClinicalMetadata(Request());
            var document = AriaReportDocumentBuilder.Build(request); document["id"] = "doc1";
            using (var harness = new Harness(document))
                TestAssert.Equal(AriaVerificationState.Verified, harness.Client.VerifyAsync("DocumentReference/doc1", request, CancellationToken.None).GetAwaiter().GetResult().State);
            foreach (Action<JObject> mutate in new Action<JObject>[] {
                d => d.Remove("author"), d => d["author"][0]["reference"] = "Practitioner/service-account",
                d => ((JArray)d["author"]).Add(new JObject { ["reference"] = "Practitioner/service-account" }),
                d => ((JArray)d["extension"]).RemoveAt(1), d => d["extension"][1]["valueString"] = "PQM_Other plan",
                d => ((JArray)d["extension"]).Add(d["extension"][1].DeepClone()) })
            {
                var altered = (JObject)document.DeepClone(); mutate(altered);
                using (var harness = new Harness(altered))
                    TestAssert.Equal(AriaVerificationState.Mismatch, harness.Client.VerifyAsync("DocumentReference/doc1", request, CancellationToken.None).GetAwaiter().GetResult().State);
            }
        }

        public static void UnboundLegacyRequestCannotBeUploaded()
        {
            using (var harness = new Harness())
            {
                var result = harness.Client.UploadAsync(Request(false), CancellationToken.None).GetAwaiter().GetResult();
                TestAssert.Equal(AriaUploadState.Rejected, result.State);
                TestAssert.Equal(0, harness.TokenCalls, "An unbound author must fail before OAuth and before any POST.");
                TestAssert.Equal(0, harness.Handler.Uris.Count);
            }
        }

        private static object ResolveAuthor(AriaFhirClient client, string userId)
        {
            var method = typeof(AriaFhirClient).GetMethod("ResolveAuthorAsync");
            TestAssert.NotNull(method, "An exact interactive-user to Practitioner resolver is required.");
            try
            {
                var task = (Task)method.Invoke(client, new object[] { userId, CancellationToken.None });
                task.GetAwaiter().GetResult();
                return task.GetType().GetProperty("Result").GetValue(task);
            }
            catch (System.Reflection.TargetInvocationException error) { throw error.InnerException; }
        }

        private static JObject Practitioner(string id, string userId, string system, bool active = true)
        { return new JObject { ["resourceType"] = "Practitioner", ["id"] = id, ["active"] = active,
            ["identifier"] = new JArray(new JObject { ["system"] = system, ["value"] = userId }) }; }

        public static void Contract()
        {
            var type = typeof(AriaReportDocumentBuilder).Assembly.GetType("ClearPlan.Core.Integration.AriaFhirClient");
            TestAssert.NotNull(type, "A bounded FHIR adapter must resolve and submit the detached report.");
            foreach (string name in new[] { "ResolvePatientAsync", "ResolveProviderAsync", "ResolveDocumentTypeAsync", "UploadAsync", "VerifyAsync" })
                TestAssert.NotNull(type.GetMethod(name), "Missing asynchronous FHIR operation: " + name);
        }

        public static void PatientExact()
        {
            using (var harness = new Harness(Bundle(Patient("p1", "TEST &+=1", "urn:synthetic"))))
            {
                var patient = harness.Client.ResolvePatientAsync("TEST &+=1", "urn:synthetic", CancellationToken.None).GetAwaiter().GetResult();
                TestAssert.Equal("Patient/p1", patient.Reference);
                TestAssert.Equal("TEST &+=1", patient.Identifier);
                TestAssert.Equal("?identifier=" + Uri.EscapeDataString("urn:synthetic|TEST &+=1"), harness.Handler.Uris[0].Query);
                TestAssert.Equal("Bearer SYNTHETIC_TOKEN", harness.Handler.Authorizations[0]);
            }
            foreach (var bundle in new[] { Bundle(), Bundle(Patient("p1", "OTHER")),
                Bundle(Patient("p1", "TEST"), Patient("p2", "TEST")), Bundle(Patient("p1", "test")) })
                using (var harness = new Harness(bundle))
                    Reject(() => harness.Client.ResolvePatientAsync("TEST", null, CancellationToken.None).GetAwaiter().GetResult());
            using (var harness = new Harness(Bundle(Patient("p1", "TEST", "urn:wrong"))))
                Reject(() => harness.Client.ResolvePatientAsync("TEST", "urn:expected", CancellationToken.None).GetAwaiter().GetResult());
        }

        public static void PatientPagination()
        {
            var first = Bundle(Patient("p1", "TEST"));
            first["link"] = new JArray(new JObject { ["relation"] = "next", ["url"] = "Patient?cursor=2" });
            using (var harness = new Harness(first, Bundle(Patient("p1", "TEST"))))
            {
                TestAssert.Equal("Patient/p1", harness.Client.ResolvePatientAsync("TEST", null, CancellationToken.None).GetAwaiter().GetResult().Reference);
                TestAssert.Equal(2, harness.Handler.Uris.Count);
            }
            using (var harness = new Harness(first, Bundle(Patient("p2", "TEST"))))
                Reject(() => harness.Client.ResolvePatientAsync("TEST", null, CancellationToken.None).GetAwaiter().GetResult());
        }

        public static void Provider()
        {
            using (var harness = new Harness(Bundle(Organization("provider1"))))
            {
                var provider = harness.Client.ResolveProviderAsync(null, CancellationToken.None).GetAwaiter().GetResult();
                TestAssert.Equal("Organization/provider1", provider.Reference);
                TestAssert.Equal("Synthetic provider", provider.Display);
                TestAssert.Equal("?type=prov&active=true", harness.Handler.Uris[0].Query);
            }
            using (var harness = new Harness(Organization("provider1")))
                TestAssert.Equal("Organization/provider1", harness.Client.ResolveProviderAsync("Organization/provider1", CancellationToken.None).GetAwaiter().GetResult().Reference);
            foreach (var value in new[] { Bundle(), Bundle(Organization("p1"), Organization("p2")), Bundle(Organization("p1", false)) })
                using (var harness = new Harness(value))
                    Reject(() => harness.Client.ResolveProviderAsync(null, CancellationToken.None).GetAwaiter().GetResult());
            using (var harness = new Harness(Organization("other")))
                Reject(() => harness.Client.ResolveProviderAsync("Organization/provider1", CancellationToken.None).GetAwaiter().GetResult());
        }

        public static void Terminology()
        {
            var nested = ValueSet(new JObject { ["system"] = TypeSystem, ["contains"] = new JArray(Coding("REPORT", "Plan report", false)) });
            foreach (var response in new[] { nested, Bundle(nested) })
                using (var harness = new Harness(response))
                {
                    var type = harness.Client.ResolveDocumentTypeAsync("Organization/p1", "REPORT", "Plan report", CancellationToken.None).GetAwaiter().GetResult();
                    TestAssert.Equal(TypeSystem, type.System);
                    TestAssert.Equal("REPORT", type.Code);
                    TestAssert.True(harness.Handler.Uris[0].Query.Contains("publisher=p1"));
                }
            using (var harness = new Harness(ValueSet(Coding("REPORT", "Plan report"))))
                TestAssert.Equal("REPORT", harness.Client.ResolveDocumentTypeAsync("Organization/p1", null, "Plan report", CancellationToken.None).GetAwaiter().GetResult().Code);
            foreach (var value in new[] { ValueSet(), ValueSet(Coding("R1", "Plan report"), Coding("R2", "Plan report")),
                ValueSet(Coding("REPORT", "Plan report", false)), ValueSet(Coding("REPORT", "Other")),
                ValueSet(new JObject { ["system"] = TypeSystem, ["code"] = "REPORT", ["display"] = "Plan report", ["inactive"] = true }) })
                using (var harness = new Harness(value))
                    Reject(() => harness.Client.ResolveDocumentTypeAsync("Organization/p1", null, "Plan report", CancellationToken.None).GetAwaiter().GetResult());
            var partial = ValueSet(Coding("REPORT", "Plan report")); partial["expansion"]["total"] = 2;
            using (var harness = new Harness(partial))
                Reject(() => harness.Client.ResolveDocumentTypeAsync("Organization/p1", "REPORT", null, CancellationToken.None).GetAwaiter().GetResult());
        }

        public static void UriBoundary()
        {
            foreach (string url in new[] { "https://other.invalid/fhir/r4/Patient?next=2", "http://example.invalid/fhir/r4/Patient",
                "https://example.invalid/other/Patient", "https://example.invalid/fhir/r40/Patient", "https://example.invalid:444/fhir/r4/Patient",
                "https://user@example.invalid/fhir/r4/Patient", "../Patient", "Patient#fragment", "Patient/%252e%252e/other", "Patient/%2fother" })
            {
                var first = Bundle(Patient("p1", "TEST"));
                first["link"] = new JArray(new JObject { ["relation"] = "next", ["url"] = url });
                using (var harness = new Harness(first))
                {
                    Reject(() => harness.Client.ResolvePatientAsync("TEST", null, CancellationToken.None).GetAwaiter().GetResult());
                    TestAssert.Equal(1, harness.Handler.Uris.Count, "No request may leave the configured service boundary.");
                    TestAssert.Equal(1, harness.TokenCalls);
                }
            }
            using (var harness = new Harness())
            {
                Reject(() => harness.Client.ResolveProviderAsync("https://other.invalid/Organization/p1", CancellationToken.None).GetAwaiter().GetResult());
                Reject(() => harness.Client.ResolvePatientAsync("TEST|other", null, CancellationToken.None).GetAwaiter().GetResult());
                TestAssert.Equal(0, harness.Handler.Uris.Count);
                TestAssert.Equal(0, harness.TokenCalls);
            }
        }

        public static void BoundedFailures()
        {
            foreach (string text in new[] { "{bad private body", new string(' ', 1048577), "{\"resourceType\":\"Bundle\",\"entry\":[],\"entry\":[{}]}" })
                using (var harness = new Harness())
                {
                    harness.Handler.Send = (r, c) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(text) });
                    Reject(() => harness.Client.ResolvePatientAsync("TEST", null, CancellationToken.None).GetAwaiter().GetResult());
                }
            using (var harness = new Harness())
            {
                harness.Handler.Send = (r, c) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest) {
                    Content = new StringContent("{\"resourceType\":\"OperationOutcome\",\"issue\":[{\"diagnostics\":\"PRIVATE SERVER TEXT\"}]}") });
                var error = Reject(() => harness.Client.ResolvePatientAsync("TEST", null, CancellationToken.None).GetAwaiter().GetResult());
                TestAssert.False(error.Message.Contains("PRIVATE"));
                TestAssert.True(error.InnerException == null);
            }
            var loop = Bundle(Patient("p1", "TEST")); loop["link"] = new JArray(new JObject { ["relation"] = "next", ["url"] = "Patient?identifier=TEST" });
            using (var harness = new Harness(loop))
            {
                Reject(() => harness.Client.ResolvePatientAsync("TEST", null, CancellationToken.None).GetAwaiter().GetResult());
                TestAssert.Equal(1, harness.Handler.Uris.Count);
            }
        }

        public static void UploadOnce()
        {
            using (var harness = new Harness())
            {
                harness.Handler.Send = async (message, token) => {
                    var payload = JObject.Parse(await message.Content.ReadAsStringAsync());
                    TestAssert.Equal("preliminary", (string)payload["docStatus"]);
                    TestAssert.Equal("Patient/p1", (string)payload["subject"]["reference"]);
                    TestAssert.Equal("POST", message.Method.Method);
                    TestAssert.Equal("application/fhir+json", message.Content.Headers.ContentType.MediaType);
                    return Response(new JObject { ["resourceType"] = "DocumentReference", ["id"] = "doc1" }, HttpStatusCode.Created);
                };
                var result = harness.Client.UploadAsync(Request(), CancellationToken.None).GetAwaiter().GetResult();
                TestAssert.Equal(AriaUploadState.Sent, result.State);
                TestAssert.Equal("DocumentReference/doc1", result.ResourceReference);
                TestAssert.Equal(1, harness.Handler.Uris.Count);
            }
            using (var harness = new Harness())
            {
                harness.Handler.Send = (r, c) => {
                    var response = new HttpResponseMessage(HttpStatusCode.Created);
                    response.Headers.Location = new Uri("https://example.invalid/fhir/r4/DocumentReference/doc2/_history/3");
                    return Task.FromResult(response);
                };
                TestAssert.Equal("DocumentReference/doc2", harness.Client.UploadAsync(Request(), CancellationToken.None).GetAwaiter().GetResult().ResourceReference);
            }
        }

        public static void UploadUncertain()
        {
            foreach (int status in new[] { 400, 401, 403, 408, 429, 500, 503, 302, 202 })
                using (var harness = new Harness())
                {
                    harness.Handler.Send = (r, c) => Task.FromResult(new HttpResponseMessage((HttpStatusCode)status) { Content = new StringContent("PRIVATE SERVER TEXT") });
                    var result = harness.Client.UploadAsync(Request(), CancellationToken.None).GetAwaiter().GetResult();
                    TestAssert.Equal(status >= 400 && status < 500 && status != 408 ? AriaUploadState.Rejected : AriaUploadState.Uncertain, result.State);
                    TestAssert.False(result.Message.Contains("PRIVATE"));
                    TestAssert.Equal(1, harness.Handler.Uris.Count);
                }
            foreach (Exception failure in new Exception[] { new HttpRequestException("PRIVATE"), new OperationCanceledException("PRIVATE") })
                using (var harness = new Harness())
                {
                    harness.Handler.Send = (r, c) => Task.FromException<HttpResponseMessage>(failure);
                    TestAssert.Equal(AriaUploadState.Uncertain, harness.Client.UploadAsync(Request(), CancellationToken.None).GetAwaiter().GetResult().State);
                    TestAssert.Equal(1, harness.Handler.Uris.Count);
                }
            using (var harness = new Harness())
            {
                harness.Handler.Send = (r, c) => {
                    var response = Response(new JObject { ["resourceType"] = "DocumentReference", ["id"] = "doc1" }, HttpStatusCode.Created);
                    response.Headers.Location = new Uri("https://other.invalid/fhir/r4/DocumentReference/doc1");
                    return Task.FromResult(response);
                };
                TestAssert.Equal(AriaUploadState.Uncertain, harness.Client.UploadAsync(Request(), CancellationToken.None).GetAwaiter().GetResult().State);
                TestAssert.Equal(1, harness.Handler.Uris.Count);
            }
        }

        public static void SafeBusinessRuleRejections()
        {
            var statusProperty = typeof(AriaUploadResult).GetProperty("HttpStatusCode");
            var codeProperty = typeof(AriaUploadResult).GetProperty("BusinessRuleCode");
            TestAssert.NotNull(statusProperty, "An upload result should retain the numeric HTTP status without a raw response.");
            TestAssert.NotNull(codeProperty, "Only explicitly allowlisted business-rule codes may reach diagnostics.");
            foreach (JObject issue in new[] {
                new JObject { ["details"] = new JObject { ["coding"] = new JArray(new JObject { ["code"] = "FUTURE_DATE_TIME" }) }, ["diagnostics"] = "PRIVATE PATIENT TEXT" },
                new JObject { ["diagnostics"] = "FUTURE_DATE_TIME: PRIVATE PATIENT TEXT" }
            })
                using (var harness = new Harness())
                {
                    harness.Handler.Send = (r, c) => Task.FromResult(Response(new JObject { ["resourceType"] = "OperationOutcome", ["issue"] = new JArray(issue) }, (HttpStatusCode)422));
                    var result = harness.Client.UploadAsync(Request(), CancellationToken.None).GetAwaiter().GetResult();
                    TestAssert.Equal(AriaUploadState.Rejected, result.State);
                    TestAssert.Equal(422, (int)statusProperty.GetValue(result));
                    TestAssert.Equal("FUTURE_DATE_TIME", (string)codeProperty.GetValue(result));
                    TestAssert.True(result.Message.Contains("FUTURE_DATE_TIME") && result.Message.Contains("DocumentDateSafetyMinutes"));
                    TestAssert.False(result.Message.Contains("PRIVATE"));
                    TestAssert.Equal(1, harness.Handler.Uris.Count);
                }
            foreach (string body in new[] { "PRIVATE FUTURE_DATE_TIME non-JSON", new string('x', 65537),
                "{\"resourceType\":\"OperationOutcome\",\"issue\":[{\"diagnostics\":\"PRIVATE_RULE_CODE\"}]}",
                "{\"resourceType\":\"OperationOutcome\",\"issue\":[{\"diagnostics\":\"PREFIX_FUTURE_DATE_TIME_SUFFIX\"}]}" })
                using (var harness = new Harness())
                {
                    harness.Handler.Send = (r, c) => Task.FromResult(new HttpResponseMessage((HttpStatusCode)422) { Content = new StringContent(body) });
                    var result = harness.Client.UploadAsync(Request(), CancellationToken.None).GetAwaiter().GetResult();
                    TestAssert.Equal(AriaUploadState.Rejected, result.State, "A bad error body must not turn a confirmed HTTP rejection into an uncertain write.");
                    TestAssert.Equal(422, (int)statusProperty.GetValue(result));
                    TestAssert.True(codeProperty.GetValue(result) == null);
                    TestAssert.False(result.Message.Contains("PRIVATE") || result.Message.Contains("PREFIX"));
                }
        }

        public static void Verification()
        {
            var request = Request(); var document = AriaReportDocumentBuilder.Build(request); document["id"] = "doc1";
            using (var harness = new Harness(document))
                TestAssert.Equal(AriaVerificationState.Verified, harness.Client.VerifyAsync("DocumentReference/doc1", request, CancellationToken.None).GetAwaiter().GetResult().State);
            foreach (Action<JObject> change in new Action<JObject>[] { j => j["subject"]["reference"] = "Patient/other", j => j["docStatus"] = "final",
                j => j["status"] = "superseded", j => j["type"]["coding"][0]["code"] = "OTHER", j => j["category"][0]["coding"][0]["code"] = "TIF",
                j => j["content"][0]["attachment"]["data"] = Convert.ToBase64String(Encoding.ASCII.GetBytes("%PDF-wrong")), j => j["id"] = "other",
                j => j["custodian"]["reference"] = "Organization/other" })
            {
                var changed = (JObject)document.DeepClone(); change(changed);
                using (var harness = new Harness(changed))
                    TestAssert.Equal(AriaVerificationState.Mismatch, harness.Client.VerifyAsync("DocumentReference/doc1", request, CancellationToken.None).GetAwaiter().GetResult().State);
            }
            var metadata = (JObject)document.DeepClone(); ((JObject)metadata["content"][0]["attachment"]).Remove("data");
            metadata["content"][0]["attachment"]["url"] = "https://external.invalid/private.pdf";
            using (var harness = new Harness(metadata))
            {
                TestAssert.Equal(AriaVerificationState.MetadataOnly, harness.Client.VerifyAsync("DocumentReference/doc1", request, CancellationToken.None).GetAwaiter().GetResult().State);
                TestAssert.Equal(1, harness.Handler.Uris.Count, "Readback must not fetch external attachment URLs.");
            }
            using (var hash = SHA1.Create()) metadata["content"][0]["attachment"]["hash"] = Convert.ToBase64String(hash.ComputeHash(request.PdfBytes));
            using (var harness = new Harness(metadata))
                TestAssert.Equal(AriaVerificationState.Verified, harness.Client.VerifyAsync("DocumentReference/doc1", request, CancellationToken.None).GetAwaiter().GetResult().State);
        }

        public static void Cancellation()
        {
            using (var harness = new Harness())
            using (var cancellation = new CancellationTokenSource())
            {
                cancellation.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => harness.Client.ResolvePatientAsync("TEST", null, cancellation.Token).GetAwaiter().GetResult());
                TestAssert.Equal(AriaUploadState.Rejected, harness.Client.UploadAsync(Request(), cancellation.Token).GetAwaiter().GetResult().State);
                TestAssert.Equal(0, harness.TokenCalls);
                TestAssert.Equal(0, harness.Handler.Uris.Count);
            }
        }

        public static void CreateReferenceRecovery()
        {
            foreach (string body in new[] { "{bad body", "{\"resourceType\":\"OperationOutcome\",\"issue\":[{\"severity\":\"information\"}]}" })
                using (var harness = new Harness())
                {
                    harness.Handler.Send = (r, c) => {
                        var response = new HttpResponseMessage(HttpStatusCode.Created) { Content = new StringContent(body) };
                        response.Headers.Location = new Uri("https://example.invalid/fhir/r4/DocumentReference/doc1/_history/2");
                        return Task.FromResult(response);
                    };
                    var result = harness.Client.UploadAsync(Request(), CancellationToken.None).GetAwaiter().GetResult();
                    TestAssert.Equal(AriaUploadState.Uncertain, result.State);
                    TestAssert.Equal("DocumentReference/doc1", result.ResourceReference, "A validated Location must survive a response-body failure for GET-only reconciliation.");
                    TestAssert.Equal(1, harness.Handler.Uris.Count);
                }
            using (var harness = new Harness())
            {
                harness.Handler.Send = (r, c) => {
                    var response = new HttpResponseMessage(HttpStatusCode.Created);
                    response.Headers.Location = new Uri("https://example.invalid/fhir/r4/DocumentReference/doc1/_history/2?_format=json");
                    return Task.FromResult(response);
                };
                var result = harness.Client.UploadAsync(Request(), CancellationToken.None).GetAwaiter().GetResult();
                TestAssert.Equal("DocumentReference/doc1", result.ResourceReference, "Only the already validated canonical resource identity is retained, never its query.");
                TestAssert.Equal(AriaUploadState.Sent, result.State);
            }
        }

        public static void ReadbackTimestampIntegrity()
        {
            var request = Request(); var document = AriaReportDocumentBuilder.Build(request); document["id"] = "doc1";
            foreach (Action<JObject> change in new Action<JObject>[] {
                j => j["date"] = request.DocumentDateUtc.AddMilliseconds(4).ToString("o"),
                j => j.Remove("date"), j => j["date"] = "09/11/2026 03:04:05",
                j => j["content"][0]["attachment"]["creation"] = request.CreatedUtc.AddMonths(2).ToString("o") })
            {
                var altered = (JObject)document.DeepClone(); change(altered);
                using (var harness = new Harness(altered))
                    TestAssert.Equal(AriaVerificationState.Mismatch, harness.Client.VerifyAsync("DocumentReference/doc1", request, CancellationToken.None).GetAwaiter().GetResult().State);
            }
            document["date"] = request.DocumentDateUtc.AddMilliseconds(-3).ToOffset(TimeSpan.FromHours(2)).ToString("o");
            document["content"][0]["attachment"]["creation"] = request.CreatedUtc.AddMilliseconds(3).ToString("o");
            using (var harness = new Harness(document))
            {
                var result = harness.Client.VerifyAsync("DocumentReference/doc1", request, CancellationToken.None).GetAwaiter().GetResult();
                TestAssert.Equal(AriaVerificationState.Verified, result.State);
                TestAssert.True((bool)typeof(AriaVerificationResult).GetProperty("AttachmentCreationChecked").GetValue(result));
            }
        }

        public static void CreatedEmptyBodyWithAriaLocation()
        {
            var stageProperty = typeof(AriaUploadResult).GetProperty("ProcessingStageCode");
            TestAssert.NotNull(stageProperty, "An uncertain create must expose its fixed processing stage without server text or URLs.");
            var previous = Thread.CurrentThread.CurrentCulture;
            try
            {
                foreach (string culture in new[] { "de-DE", "en-US" })
                foreach (string trailingSlash in new[] { "", "/" })
                {
                    Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo(culture);
                    const string reference = "DocumentReference/DocumentReference-48500000000000000999-3-5";
                    var request = Request(); var document = AriaReportDocumentBuilder.Build(request);
                    document["id"] = reference.Substring("DocumentReference/".Length);
                    using (var harness = new Harness())
                    {
                        harness.Handler.Send = (r, c) => {
                            if (r.Method == HttpMethod.Get) return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                                RequestMessage = r, Content = new StringContent(document.ToString()) });
                            var response = new HttpResponseMessage(HttpStatusCode.Created) {
                                RequestMessage = r, Content = new StreamContent(new System.IO.MemoryStream(new byte[0])) };
                            response.Content.Headers.ContentLength = 0;
                            response.Headers.Location = new Uri("https://example.invalid:55370/fhir/r4/" + reference + "/_history/1");
                            TestAssert.True(response.Content.Headers.ContentType == null);
                            return Task.FromResult(response);
                        };
                        var client = new AriaFhirClient(harness.Http, new Uri("https://example.invalid:55370/fhir/r4" + trailingSlash),
                            c => Task.FromResult("SYNTHETIC_TOKEN"), 5);
                        var result = client.UploadAsync(request, CancellationToken.None).GetAwaiter().GetResult();
                        TestAssert.Equal(AriaUploadState.Sent, result.State, "A same-origin 201 with valid Location and an empty body must be accepted.");
                        TestAssert.Equal(reference, result.ResourceReference);
                        TestAssert.Equal("accepted", (string)stageProperty.GetValue(result));
                        TestAssert.Equal(AriaVerificationState.Verified, client.VerifyAsync(result.ResourceReference, request, CancellationToken.None).GetAwaiter().GetResult().State);
                        TestAssert.Equal(2, harness.Handler.Uris.Count, "Create must be followed by one readback, never a repeated POST.");
                    }
                }
            }
            finally { Thread.CurrentThread.CurrentCulture = previous; }
            foreach (string stage in new[] { "response-target", "create-location", "create-body", "create-reference" })
                using (var harness = new Harness())
                {
                    harness.Handler.Send = (r, c) => {
                        var response = new HttpResponseMessage(HttpStatusCode.Created) {
                            RequestMessage = new HttpRequestMessage(HttpMethod.Post, stage == "response-target" ? "https://outside.invalid/PRIVATE" : r.RequestUri.AbsoluteUri),
                            Content = stage == "create-body" ? new StringContent("{PRIVATE malformed JSON") :
                                stage == "create-reference" ? new StringContent("{\"resourceType\":\"DocumentReference\",\"id\":\"different\"}") : null };
                        response.Headers.Location = new Uri(stage == "create-location" ? "https://outside.invalid/PRIVATE" : "https://example.invalid/fhir/r4/DocumentReference/doc1/_history/1");
                        return Task.FromResult(response);
                    };
                    var result = harness.Client.UploadAsync(Request(), CancellationToken.None).GetAwaiter().GetResult();
                    TestAssert.Equal(AriaUploadState.Uncertain, result.State);
                    TestAssert.Equal(stage, (string)stageProperty.GetValue(result));
                    TestAssert.Equal(stage == "create-body" || stage == "response-target" ? "DocumentReference/doc1" : null, result.ResourceReference);
                    TestAssert.False(result.Message.Contains("PRIVATE") || result.Message.Contains("https://"));
                    TestAssert.Equal(1, harness.Handler.Uris.Count);
                }
        }

        public static void UnexpectedResponseTargetRetainsOnlyTrustedRecoveryLocation()
        {
            var request = Request();
            foreach (string variant in new[] { "trusted-empty", "trusted-matching", "trusted-malformed", "foreign", "wrong-type", "missing", "conflicting" })
                using (var harness = new Harness())
                {
                    int posts = 0, gets = 0;
                    var stored = AriaReportDocumentBuilder.Build(request); stored["id"] = "doc1";
                    harness.Handler.Send = (r, c) => {
                        if (r.Method == HttpMethod.Get) { gets++; return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { RequestMessage = r, Content = new StringContent(stored.ToString()) }); }
                        posts++;
                        var response = new HttpResponseMessage(HttpStatusCode.Created) {
                            RequestMessage = new HttpRequestMessage(HttpMethod.Post, "https://example.invalid/fhir/r4/DocumentReference/unexpected"),
                            Content = variant == "trusted-malformed" ? new StringContent("{invalid") :
                                variant == "trusted-matching" || variant == "missing" ? new StringContent(stored.ToString()) :
                                variant == "conflicting" ? new StringContent("{\"resourceType\":\"DocumentReference\",\"id\":\"different\"}") : null };
                        if (variant != "missing") response.Headers.Location = new Uri(variant == "foreign" ? "https://outside.invalid/fhir/r4/DocumentReference/doc1" :
                            variant == "wrong-type" ? "https://example.invalid/fhir/r4/Patient/doc1" : "https://example.invalid/fhir/r4/DocumentReference/doc1/_history/1");
                        return Task.FromResult(response);
                    };
                    var result = harness.Client.UploadAsync(request, CancellationToken.None).GetAwaiter().GetResult();
                    TestAssert.Equal(AriaUploadState.Uncertain, result.State, "An unexpected response target must never be reported as Sent.");
                    bool recoverable = variant.StartsWith("trusted-", StringComparison.Ordinal);
                    TestAssert.Equal(recoverable ? "DocumentReference/doc1" : null, result.ResourceReference,
                        "Only a independently validated, nonconflicting Location can authorize GET-only reconciliation; body-only references are insufficient.");
                    TestAssert.Equal(1, posts); TestAssert.Equal(0, gets, "Create itself must not imply completed readback.");
                    if (recoverable)
                    {
                        TestAssert.Equal(AriaVerificationState.Verified, harness.Client.VerifyAsync(result.ResourceReference, request, CancellationToken.None).GetAwaiter().GetResult().State);
                        stored["subject"]["reference"] = "Patient/other";
                        TestAssert.Equal(AriaVerificationState.Mismatch, harness.Client.VerifyAsync(result.ResourceReference, request, CancellationToken.None).GetAwaiter().GetResult().State,
                            "A recovery candidate is not sufficient proof: patient and PDF checks remain mandatory.");
                        TestAssert.Equal(1, posts); TestAssert.Equal(2, gets);
                    }
                }
        }

        public static void SparseFileReadback()
        {
            var request = DescribedRequest(); var sparse = SparseDocument(request);
            int fileReads = 0;
            using (var harness = new Harness(sparse))
            {
                var client = WithAttachmentReader(harness, (path, token) => { fileReads++; TestAssert.Equal("\\\\synthetic.invalid\\documents\\test.pdf", path); return Task.FromResult(request.PdfBytes); });
                var result = client.VerifyAsync("DocumentReference/doc1", request, CancellationToken.None).GetAwaiter().GetResult();
                TestAssert.Equal(AriaVerificationState.Verified, result.State);
                TestAssert.Equal(1, fileReads);
                TestAssert.False((bool)typeof(AriaVerificationResult).GetProperty("AttachmentCreationChecked").GetValue(result));
                TestAssert.True(result.Message.Contains("creation") && result.Message.Contains("not checked"));
            }
            using (var harness = new Harness(sparse))
                TestAssert.Equal(AriaVerificationState.MetadataOnly, harness.Client.VerifyAsync("DocumentReference/doc1", request, CancellationToken.None).GetAwaiter().GetResult().State);
            foreach (Action<JObject> change in new Action<JObject>[] { j => j["description"] = "Other description",
                j => j["content"][0]["attachment"]["title"] = "wrong.pdf", j => j["category"][0]["coding"][0]["code"] = "PDF" })
            {
                var changed = (JObject)sparse.DeepClone(); change(changed);
                using (var harness = new Harness(changed))
                {
                    int reads = 0;
                    var result = WithAttachmentReader(harness, (p, c) => { reads++; return Task.FromResult(request.PdfBytes); })
                        .VerifyAsync("DocumentReference/doc1", request, CancellationToken.None).GetAwaiter().GetResult();
                    TestAssert.Equal(AriaVerificationState.Mismatch, result.State);
                    TestAssert.Equal(0, reads, "Mismatched metadata must not trigger a filesystem read.");
                }
            }
            using (var harness = new Harness(sparse))
                TestAssert.Equal(AriaVerificationState.Mismatch, WithAttachmentReader(harness, (p, c) => Task.FromResult(Encoding.ASCII.GetBytes("%PDF-wrong")))
                    .VerifyAsync("DocumentReference/doc1", request, CancellationToken.None).GetAwaiter().GetResult().State);
            sparse["content"][0]["attachment"]["url"] = "https://external.invalid/test.pdf";
            using (var harness = new Harness(sparse))
            {
                var result = WithAttachmentReader(harness, (p, c) => { throw new Exception("HTTP attachments must never be followed."); })
                    .VerifyAsync("DocumentReference/doc1", request, CancellationToken.None).GetAwaiter().GetResult();
                TestAssert.Equal(AriaVerificationState.MetadataOnly, result.State);
            }
        }

        private static AriaReportUploadRequest DescribedRequest()
        { var r = Request(); return AriaReportUploadTests.BindClinicalMetadata(new AriaReportUploadRequest(r.ReportPatientId, r.ResolvedPatientId, r.PatientReference, r.ActivePlanKey,
            r.OrganizationReference, r.DocumentTypeSystem, r.DocumentTypeCode, r.DocumentTypeDisplay, r.PdfBytes, r.Title, r.CreatedUtc,
            "Synthetic readback description", "TIF", "TIF")); }
        private static JObject SparseDocument(AriaReportUploadRequest request)
        { var document = AriaReportDocumentBuilder.Build(request); document["id"] = "doc1";
            document["content"] = new JArray(new JObject { ["attachment"] = new JObject { ["id"] = "synthetic-attachment", ["contentType"] = "application/pdf", ["url"] = "\\\\synthetic.invalid\\documents\\test.pdf" } });
            return document; }
        private static AriaFhirClient WithAttachmentReader(Harness harness, Func<string, CancellationToken, Task<byte[]>> reader)
        {
            var constructor = typeof(AriaFhirClient).GetConstructors().SingleOrDefault(c => c.GetParameters().Length == 5);
            TestAssert.NotNull(constructor, "A separately configured attachment reader must be an explicit optional dependency.");
            return (AriaFhirClient)constructor.Invoke(new object[] { harness.Http, new Uri("https://example.invalid/fhir/r4"),
                new Func<CancellationToken, Task<string>>(c => Task.FromResult("SYNTHETIC_TOKEN")), 45, reader });
        }

        public static void SlowDependencyCancellation()
        {
            using (var cancellation = new CancellationTokenSource())
            using (var http = new HttpClient(new FakeHandler()))
            {
                var never = new TaskCompletionSource<string>();
                var client = new AriaFhirClient(http, new Uri("https://example.invalid/fhir/r4"), token => never.Task);
                var operation = client.ResolvePatientAsync("TEST", null, cancellation.Token);
                cancellation.Cancel();
                TestAssert.True(Task.WhenAny(operation, Task.Delay(1500)).GetAwaiter().GetResult() == operation,
                    "Cancellation must bound a token provider even when it ignores its token.");
                TestAssert.Throws<OperationCanceledException>(() => operation.GetAwaiter().GetResult());
            }
            using (var harness = new Harness())
            using (var cancellation = new CancellationTokenSource())
            {
                harness.Handler.Send = (r, c) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new NeverReadyContent() });
                var operation = harness.Client.ResolvePatientAsync("TEST", null, cancellation.Token);
                cancellation.Cancel();
                TestAssert.True(Task.WhenAny(operation, Task.Delay(1500)).GetAwaiter().GetResult() == operation,
                    "Cancellation must bound content stream acquisition as well as network headers.");
                TestAssert.Throws<OperationCanceledException>(() => operation.GetAwaiter().GetResult());
            }
        }

        public static void TerminologyIntegrity()
        {
            using (var harness = new Harness(ValueSet(Coding("REPORT", "Plan report"), Coding("REPORT", "Conflicting display"))))
                Reject(() => harness.Client.ResolveDocumentTypeAsync("Organization/p1", "REPORT", "Plan report", CancellationToken.None).GetAwaiter().GetResult());
            var skipped = ValueSet(Coding("REPORT", "Plan report")); skipped["expansion"]["offset"] = 1;
            using (var harness = new Harness(skipped))
                Reject(() => harness.Client.ResolveDocumentTypeAsync("Organization/p1", "REPORT", null, CancellationToken.None).GetAwaiter().GetResult());
            var first = Bundle(ValueSet(Coding("OTHER", "Other")));
            first["entry"][0]["resource"]["expansion"]["total"] = 2;
            first["link"] = new JArray(new JObject { ["relation"] = "next", ["url"] = "ValueSet/$expand?cursor=2" });
            var second = ValueSet(Coding("REPORT", "Plan report")); second["expansion"]["offset"] = 1; second["expansion"]["total"] = 2;
            using (var harness = new Harness(first, second))
            {
                TestAssert.Equal("REPORT", harness.Client.ResolveDocumentTypeAsync("Organization/p1", "REPORT", null, CancellationToken.None).GetAwaiter().GetResult().Code);
                TestAssert.Equal(2, harness.Handler.Uris.Count);
            }
        }

        public static void PageLimitAndUploadPreflight()
        {
            using (var harness = new Harness())
            {
                int page = 0;
                harness.Handler.Send = (r, c) => {
                    var value = Bundle(Patient("p1", "TEST"));
                    value["link"] = new JArray(new JObject { ["relation"] = "next", ["url"] = "Patient?cursor=" + (++page) });
                    return Task.FromResult(Response(value));
                };
                Reject(() => harness.Client.ResolvePatientAsync("TEST", null, CancellationToken.None).GetAwaiter().GetResult());
                TestAssert.Equal(20, harness.Handler.Uris.Count);
            }
            var handler = new FakeHandler();
            using (var http = new HttpClient(handler))
            {
                var client = new AriaFhirClient(http, new Uri("https://example.invalid/fhir/r4"), token => Task.FromException<string>(new Exception("PRIVATE SECRET")));
                var result = client.UploadAsync(Request(), CancellationToken.None).GetAwaiter().GetResult();
                TestAssert.Equal(AriaUploadState.Rejected, result.State);
                TestAssert.False(result.Message.Contains("PRIVATE"));
                TestAssert.Equal(0, handler.Uris.Count);
            }
        }

        private const string TypeSystem = "http://varian.com/fhir/CodeSystem/DocumentReference/documentreference-type";
        private static AriaReportUploadRequest Request(bool bind = true)
        { var request = new AriaReportUploadRequest("SYNTHETIC", "SYNTHETIC", "Patient/p1", "synthetic-plan", "Organization/p1", TypeSystem,
            "REPORT", "Plan report", Encoding.ASCII.GetBytes("%PDF-synthetic-only"), "Synthetic review.pdf", new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero));
            return bind ? AriaReportUploadTests.BindClinicalMetadata(request) : request; }
        private static JObject Patient(string id, string value, string system = null)
        { return new JObject { ["resourceType"] = "Patient", ["id"] = id, ["identifier"] = new JArray(new JObject { ["system"] = system, ["value"] = value }) }; }
        private static JObject Organization(string id, bool active = true)
        { return new JObject { ["resourceType"] = "Organization", ["id"] = id, ["active"] = active, ["name"] = "Synthetic provider",
            ["type"] = new JArray(new JObject { ["coding"] = new JArray(new JObject { ["system"] = "http://terminology.hl7.org/CodeSystem/organization-type", ["code"] = "prov" }) }) }; }
        private static JObject Coding(string code, string display, bool withSystem = true)
        { var value = new JObject { ["code"] = code, ["display"] = display }; if (withSystem) value["system"] = TypeSystem; return value; }
        private static JObject ValueSet(params JObject[] codes)
        { return new JObject { ["resourceType"] = "ValueSet", ["url"] = "http://varian.com/fhir/ValueSet/documentreference-type", ["expansion"] = new JObject { ["contains"] = new JArray(codes) } }; }
        private static JObject Bundle(params JObject[] resources)
        { return new JObject { ["resourceType"] = "Bundle", ["type"] = "searchset", ["entry"] = new JArray(resources.Select(r => new JObject { ["resource"] = r })) }; }
        private static HttpResponseMessage Response(JObject resource, HttpStatusCode status = HttpStatusCode.OK)
        { return new HttpResponseMessage(status) { Content = new StringContent(resource.ToString(), Encoding.UTF8, "application/fhir+json") }; }
        private static AriaFhirException Reject(Action action)
        { return TestAssert.Throws<AriaFhirException>(action); }

        private sealed class Harness : IDisposable
        {
            private readonly HttpClient http;
            public HttpClient Http { get { return http; } }
            public readonly FakeHandler Handler = new FakeHandler();
            public readonly AriaFhirClient Client;
            public int TokenCalls;
            public Harness(params JObject[] responses)
            {
                var queue = new Queue<JObject>(responses);
                Handler.Send = (r, c) => Task.FromResult(Response(queue.Dequeue()));
                http = new HttpClient(Handler);
                Client = new AriaFhirClient(http, new Uri("https://example.invalid/fhir/r4"), token => { TokenCalls++; return Task.FromResult("SYNTHETIC_TOKEN"); });
            }
            public void Dispose() { http.Dispose(); }
        }
        private sealed class FakeHandler : HttpMessageHandler
        {
            public Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> Send;
            public readonly List<Uri> Uris = new List<Uri>();
            public readonly List<string> Authorizations = new List<string>();
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellation)
            { Uris.Add(request.RequestUri); Authorizations.Add(request.Headers.Authorization == null ? "" : request.Headers.Authorization.ToString()); return Send(request, cancellation); }
        }
        private sealed class NeverReadyContent : HttpContent
        {
            protected override Task SerializeToStreamAsync(System.IO.Stream stream, TransportContext context)
            { return new TaskCompletionSource<bool>().Task; }
            protected override bool TryComputeLength(out long length) { length = 0; return false; }
        }
    }
}
