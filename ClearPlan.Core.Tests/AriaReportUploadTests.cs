using System;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using ClearPlan.Core.Integration;
using ClearPlan.Core.Review;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ClearPlan.Core.Tests
{
    internal static class AriaReportUploadTests
    {
        public static void ExposesDetachedBuilderContract()
        {
            var builder = typeof(ReviewSnapshot).Assembly.GetType(
                "ClearPlan.Core.Integration.AriaReportDocumentBuilder");
            TestAssert.NotNull(builder, "A detached ARIA document builder is required.");
        }

        public static void PayloadRoundTripsExactBytesAndFingerprint()
        {
            byte[] bytes = Pdf();
            var request = Request(pdfBytes: bytes);
            JObject payload = AriaReportDocumentBuilder.Build(request);
            TestAssert.NotNull(payload, "A PDF DocumentReference payload is required.");
            var attachment = payload["content"][0]["attachment"];
            TestAssert.True(bytes.SequenceEqual(Convert.FromBase64String((string)attachment["data"])));
            TestAssert.Equal("application/pdf", (string)attachment["contentType"]);
            TestAssert.Equal("Patient/synthetic-fhir-1", (string)payload["subject"]["reference"]);
            TestAssert.Equal("Organization/synthetic-provider-1", (string)payload["custodian"]["reference"]);
            using (var sha = SHA256.Create())
            {
                TestAssert.Equal(BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant(),
                    request.PdfSha256);
            }
            TestAssert.Equal(bytes.Length, request.PdfLength);
            TestAssert.Equal("synthetic-course|synthetic-plan", request.ActivePlanKey);
            TestAssert.Equal("SYNTHETIC-001", request.ReportPatientId);
            TestAssert.Equal(request.ReportPatientId, request.ResolvedPatientId);
            TestAssert.True(attachment["hash"] == null, "FHIR R4 attachment.hash is not a SHA256 field.");
        }

        public static void RequestAndPayloadMutationsCannotChangePreparedBytes()
        {
            byte[] bytes = Pdf();
            byte[] original = (byte[])bytes.Clone();
            var request = Request(pdfBytes: bytes);
            bytes[0] = 0;
            byte[] exposed = request.PdfBytes;
            TestAssert.True(original.SequenceEqual(exposed), "The input PDF must be defensively copied.");
            exposed[1] = 0;
            TestAssert.True(original.SequenceEqual(request.PdfBytes), "The PDF getter must return a defensive copy.");
            TestAssert.True(typeof(AriaReportUploadRequest).IsSealed);
            TestAssert.True(typeof(AriaReportUploadRequest).GetProperties().All(p => p.GetSetMethod() == null));
            var first = AriaReportDocumentBuilder.Build(request);
            first["subject"]["reference"] = "Patient/synthetic-other";
            first["content"][0]["attachment"]["data"] = "changed";
            var second = AriaReportDocumentBuilder.Build(request);
            TestAssert.Equal("Patient/synthetic-fhir-1", (string)second["subject"]["reference"]);
            TestAssert.True(original.SequenceEqual(Convert.FromBase64String((string)second["content"][0]["attachment"]["data"])));
        }

        public static void ExactPatientBindingHasNoNormalizationOrBypass()
        {
            foreach (string target in new[] { "SYNTHETIC-002", "synthetic-001", "SYNTHETIC-001 ", " SYNTHETIC-001" })
            {
                Reject(() => Request(resolvedPatientId: target));
            }
            foreach (string invalid in new[] { null, "", " ", "SYNTHETIC\r\n001", new string('a', 257) })
            {
                Reject(() => Request(reportPatientId: invalid, resolvedPatientId: invalid));
            }
            TestAssert.Equal("SYNTHETIC-001", Request().ReportPatientId);
        }

        public static void ReferencesRejectAbsoluteQueriesControlsAndTraversal()
        {
            foreach (string invalid in new[]
            {
                null, "", "Patient/", "patient/abc", "Patient/a/b", "Patient/../abc",
                "Patient/%2e%2e", "https://example.invalid/Patient/abc", "Patient/abc?x=1",
                "Patient/abc#fragment", "Patient/abc\r\nInjected: x", "Patient/abc\n",
                "Patient/abc\"},\"status\":\"final", " Patient/abc", "Patient/abc ",
                "Patient/" + new string('a', 65), "Patient/\u202eabc", "Patient/\\abc"
            }) Reject(() => Request(patientReference: invalid));
            foreach (string invalid in new[] { null, "", "Patient/abc", "Organization/a/b", "Organization/a\n", "Organization/" + new string('a', 65) })
                Reject(() => Request(organizationReference: invalid));
            TestAssert.Equal("Patient/a.B-1", Request(patientReference: "Patient/a.B-1").PatientReference);
            TestAssert.Equal("Organization/" + new string('a', 64),
                Request(organizationReference: "Organization/" + new string('a', 64)).OrganizationReference);
        }

        public static void PdfGateRejectsMissingWrongAndOversizedData()
        {
            Reject(() => new AriaReportUploadRequest("SYNTHETIC-001", "SYNTHETIC-001", "Patient/synthetic-1",
                "synthetic-plan", "Organization/synthetic-1", "urn:oid:1.2.3.4", "synthetic-code", "Synthetic report",
                null, "Review.pdf", DateTimeOffset.UtcNow));
            foreach (byte[] invalid in new[] { new byte[0], Encoding.ASCII.GetBytes("%PDF"), Encoding.ASCII.GetBytes("%pdf-"), Encoding.ASCII.GetBytes("x%PDF-"), new byte[20 * 1024 * 1024 + 1] })
                Reject(() => Request(pdfBytes: invalid));
            TestAssert.Equal(5, Request(pdfBytes: Encoding.ASCII.GetBytes("%PDF-")).PdfLength,
                "The signature is a format gate, not a PDF parser or patient-text verifier.");
            var boundary = new byte[20 * 1024 * 1024];
            Array.Copy(Encoding.ASCII.GetBytes("%PDF-"), boundary, 5);
            TestAssert.Equal(boundary.Length, Request(pdfBytes: boundary).PdfLength);
        }

        public static void PayloadUsesInstalledPreliminaryProfileWithoutInventedIdentity()
        {
            JObject payload = AriaReportDocumentBuilder.Build(Request());
            TestAssert.NotNull(payload, "A preliminary DocumentReference is required.");
            TestAssert.Equal("DocumentReference", (string)payload["resourceType"]);
            TestAssert.Equal("current", (string)payload["status"]);
            TestAssert.Equal("preliminary", (string)payload["docStatus"]);
            TestAssert.Equal("http://varian.com/fhir/v1/StructureDefinition/DocumentReference", (string)payload["meta"]["profile"][0]);
            TestAssert.True(payload["id"] == null && payload["identifier"] == null && payload["authenticator"] == null);
            TestAssert.Equal("http://varian.com/fhir/v1/StructureDefinition/documentreference-documentLocation", (string)payload["extension"][0]["url"]);
            TestAssert.Equal("file-server", (string)payload["extension"][0]["valueString"]);
            TestAssert.Equal(1, payload["extension"].Count());
            TestAssert.Equal(1, payload["content"].Count());
            TestAssert.Equal("http://varian.com/fhir/CodeSystem/DocumentReference/documentreference-class", (string)payload["category"][0]["coding"][0]["system"]);
            TestAssert.Equal("PDF", (string)payload["category"][0]["coding"][0]["code"]);
            TestAssert.Equal("PDF", (string)payload["category"][0]["coding"][0]["display"]);
            TestAssert.Equal("urn:oid:1.2.3.4", (string)payload["type"]["coding"][0]["system"]);
            TestAssert.Equal("synthetic-code", (string)payload["type"]["coding"][0]["code"]);
            TestAssert.Equal("Synthetic report", (string)payload["type"]["coding"][0]["display"]);
            TestAssert.True(payload["description"] == null);
        }

        public static void TimestampIsCallerSuppliedUtcInstant()
        {
            var time = new DateTimeOffset(2026, 7, 27, 14, 15, 16, 789, TimeSpan.FromHours(2));
            var request = Request(createdUtc: time);
            TestAssert.Equal(TimeSpan.Zero, request.CreatedUtc.Offset);
            TestAssert.Equal(time.ToUniversalTime(), request.CreatedUtc);
            JObject payload = AriaReportDocumentBuilder.Build(request);
            string date = (string)payload["date"];
            TestAssert.True(date.EndsWith("Z", StringComparison.Ordinal));
            TestAssert.Equal(date, (string)payload["content"][0]["attachment"]["creation"]);
            TestAssert.Equal(time, DateTimeOffset.Parse(date, CultureInfo.InvariantCulture));
            TestAssert.Equal(payload.ToString(Formatting.None), AriaReportDocumentBuilder.Build(request).ToString(Formatting.None));
            Reject(() => Request(createdUtc: default(DateTimeOffset)));
        }

        public static void DocumentDateConfiguration()
        {
            var property = typeof(AriaUploadConfiguration).GetProperty("DocumentDateSafetyMinutes");
            TestAssert.NotNull(property, "A site may explicitly configure document-date safety without changing PDF creation.");
            TestAssert.Equal(0, (int)property.GetValue(AriaUploadConfiguration.Parse("{\"Enabled\":false}")));
            foreach (int value in new[] { 0, 10, 60 })
                TestAssert.Equal(value, (int)property.GetValue(AriaUploadConfiguration.Parse("{\"Enabled\":false,\"DocumentDateSafetyMinutes\":" + value + "}")));
            foreach (string invalid in new[] { "-1", "61", "1.5", "null", "\"ten\"" })
                TestAssert.Throws<FormatException>(() => AriaUploadConfiguration.Parse("{\"Enabled\":false,\"DocumentDateSafetyMinutes\":" + invalid + "}"));
        }

        public static void DocumentDatePolicyIsPureAndBounded()
        {
            var created = new DateTimeOffset(2026, 9, 11, 12, 0, 0, TimeSpan.FromHours(2));
            var now = created.AddMinutes(3);
            TestAssert.Equal(created.ToUniversalTime(), CalculateDocumentDate(created, now, 0));
            TestAssert.Equal(now.AddMinutes(-10).ToUniversalTime(), CalculateDocumentDate(created, now, 10));
            TestAssert.Equal(created.AddMinutes(-30).ToUniversalTime(), CalculateDocumentDate(created.AddMinutes(-30), now, 10));
            TestAssert.Equal(now.ToUniversalTime(), CalculateDocumentDate(created.AddHours(1), now, 0));
            TestAssert.Equal(TimeSpan.Zero, CalculateDocumentDate(created, now, 60).Offset);
            foreach (int invalid in new[] { -1, 61 })
                TestAssert.Throws<ArgumentException>(() => CalculateDocumentDate(created, now, invalid));
            TestAssert.Throws<ArgumentException>(() => CalculateDocumentDate(default(DateTimeOffset), now, 0));
            TestAssert.Throws<ArgumentException>(() => CalculateDocumentDate(created, default(DateTimeOffset), 0));
            TestAssert.Throws<ArgumentException>(() => CalculateDocumentDate(created, DateTimeOffset.MinValue.AddMinutes(1), 60));
        }

        public static void DocumentDateNeverChangesAttachmentCreation()
        {
            var baseline = Request();
            var legacyConstructor = typeof(AriaReportUploadRequest).GetConstructors().SingleOrDefault(c => c.GetParameters().Length == 14);
            TestAssert.NotNull(legacyConstructor, "The original 14-parameter constructor must remain callable by existing binaries.");
            TestAssert.True(legacyConstructor.GetParameters().Skip(11).All(p => p.IsOptional), "The original source-level optional arguments must also remain available.");
            var legacyRequest = (AriaReportUploadRequest)legacyConstructor.Invoke(new object[] { baseline.ReportPatientId, baseline.ResolvedPatientId,
                baseline.PatientReference, baseline.ActivePlanKey, baseline.OrganizationReference, baseline.DocumentTypeSystem,
                baseline.DocumentTypeCode, baseline.DocumentTypeDisplay, baseline.PdfBytes, baseline.Title, baseline.CreatedUtc,
                baseline.Description, baseline.CategoryCode, baseline.CategoryDisplay });
            TestAssert.Equal(AriaReportDocumentBuilder.Build(baseline).ToString(Formatting.None),
                AriaReportDocumentBuilder.Build(legacyRequest).ToString(Formatting.None));
            var separate = baseline.CreatedUtc.AddMinutes(-10);
            var constructor = typeof(AriaReportUploadRequest).GetConstructors().SingleOrDefault(c => c.GetParameters().Length == 15);
            TestAssert.NotNull(constructor, "Prepared requests need an optional distinct document date.");
            Func<DateTimeOffset?, AriaReportUploadRequest> dated = value => {
                try { return (AriaReportUploadRequest)constructor.Invoke(new object[] { baseline.ReportPatientId, baseline.ResolvedPatientId,
                    baseline.PatientReference, baseline.ActivePlanKey, baseline.OrganizationReference, baseline.DocumentTypeSystem,
                    baseline.DocumentTypeCode, baseline.DocumentTypeDisplay, baseline.PdfBytes, baseline.Title, baseline.CreatedUtc,
                    baseline.Description, baseline.CategoryCode, baseline.CategoryDisplay, value }); }
                catch (System.Reflection.TargetInvocationException error) { throw error.InnerException; }
            };
            var request = dated(separate);
            var payload = AriaReportDocumentBuilder.Build(request);
            TestAssert.Equal(separate, DateTimeOffset.Parse((string)payload["date"], CultureInfo.InvariantCulture));
            TestAssert.Equal(baseline.CreatedUtc, DateTimeOffset.Parse((string)payload["content"][0]["attachment"]["creation"], CultureInfo.InvariantCulture));
            TestAssert.Equal(baseline.PdfSha256, request.PdfSha256);
            TestAssert.True(baseline.PdfBytes.SequenceEqual(request.PdfBytes));
            var property = typeof(AriaReportUploadRequest).GetProperty("DocumentDateUtc");
            TestAssert.Equal(separate, (DateTimeOffset)property.GetValue(request));
            TestAssert.Equal(baseline.CreatedUtc, (DateTimeOffset)property.GetValue(dated(null)));
            TestAssert.Throws<ArgumentException>(() => dated(baseline.CreatedUtc.AddSeconds(1)));
            TestAssert.Throws<ArgumentException>(() => dated(default(DateTimeOffset)));
        }

        public static void NativeDocumentDateDisclosure()
        {
            string source = System.IO.File.ReadAllText(System.IO.Path.Combine("ClearPlan.Script", "MainView.AriaUpload.cs"));
            TestAssert.True(source.Contains("AriaDocumentDatePolicy.Calculate") && source.Contains("config.DocumentDateSafetyMinutes"));
            TestAssert.True(source.Contains("document.GeneratedUtc") && source.Contains("ARIA-Dokumentdatum (UTC)") && source.Contains("Bericht erstellt (UTC)"));
            TestAssert.True(source.Contains("request.DocumentDateUtc != request.CreatedUtc"), "A compatibility date difference must be disclosed before send.");
            var example = JObject.Parse(System.IO.File.ReadAllText(System.IO.Path.Combine("ClearPlan.Script", "Distribution", "AriaUpload.example.json")));
            TestAssert.Equal(0, (int)example["DocumentDateSafetyMinutes"]);
        }

        private static DateTimeOffset CalculateDocumentDate(DateTimeOffset created, DateTimeOffset now, int minutes)
        {
            var type = typeof(AriaReportDocumentBuilder).Assembly.GetType("ClearPlan.Core.Integration.AriaDocumentDatePolicy");
            TestAssert.NotNull(type, "Document-date handling must be a pure explicit policy.");
            try { return (DateTimeOffset)type.GetMethod("Calculate").Invoke(null, new object[] { created, now, minutes }); }
            catch (System.Reflection.TargetInvocationException error) { throw error.InnerException; }
        }

        public static void TitleAndDescriptionAreConciseAndNeverRawPaths()
        {
            var request = Request(title: "  ClearPlan   review.PDF  ", description: "  Synthetic review only.  ");
            TestAssert.Equal("ClearPlan review.pdf", request.Title);
            TestAssert.Equal("Synthetic review only.", request.Description);
            var payload = AriaReportDocumentBuilder.Build(request);
            TestAssert.Equal(request.Title, (string)payload["content"][0]["attachment"]["title"]);
            TestAssert.Equal(request.Description, (string)payload["description"]);
            TestAssert.Equal("Review.pdf", Request(title: "Review").Title);
            foreach (string invalid in new[] { null, "", " ", ".pdf", "C:\\private\\review.pdf", "\\\\server\\share\\review.pdf", "/private/review.pdf", "../review.pdf", "review\n.pdf", "<script>.pdf", new string('a', 201) })
                Reject(() => Request(title: invalid));
            foreach (string invalid in new[] { "C:\\private\\review.pdf", "\\\\server\\share\\review.pdf", "/private/review.pdf", "review\r\nsecret", "review\u202esecret", new string('a', 513) })
                Reject(() => Request(description: invalid));
            TestAssert.True(AriaReportDocumentBuilder.Build(Request(description: "  "))["description"] == null);
        }

        public static void ResolvedTerminologyAndPlanBindingAreMandatory()
        {
            foreach (string invalid in new[] { null, "", " ", "bad\nplan", new string('a', 513) })
                Reject(() => Request(activePlanKey: invalid));
            foreach (string invalid in new[] { null, "", "relative-system", "javascript:alert(1)", "https://example.invalid/\n", "https://user:password@example.invalid/type" })
                Reject(() => Request(documentTypeSystem: invalid));
            foreach (string invalid in new[] { null, "", " ", "bad\rcode", new string('a', 257) })
            {
                Reject(() => Request(documentTypeCode: invalid));
                Reject(() => Request(documentTypeDisplay: invalid));
            }
            foreach (string invalid in new[] { null, "", "TIFX", "Patient Document", "pdf", "PDF\n" })
                Reject(() => Request(categoryCode: invalid));
            Reject(() => Request(categoryDisplay: ""));
            Reject(() => Request(categoryDisplay: "PDF\r\nInjected"));
            var tif = AriaReportDocumentBuilder.Build(Request(categoryCode: "TIF", categoryDisplay: "Image document"));
            TestAssert.Equal("TIF", (string)tif["category"][0]["coding"][0]["code"]);
            TestAssert.Equal("Image document", (string)tif["category"][0]["coding"][0]["display"]);
        }

        public static void ValidationErrorsAreSanitizedAndCarryNoInnerException()
        {
            string secret = "SYNTHETIC-SENSITIVE-MARKER";
            var mismatch = TestAssert.Throws<ArgumentException>(() => Request(resolvedPatientId: secret));
            TestAssert.Equal("The report patient does not match the resolved ARIA patient.", mismatch.Message);
            TestAssert.True(mismatch.InnerException == null);
            var reference = TestAssert.Throws<ArgumentException>(() => Request(patientReference: "Patient/" + secret + "\n"));
            TestAssert.False(reference.Message.Contains(secret));
            TestAssert.True(reference.InnerException == null);
            var missing = TestAssert.Throws<ArgumentException>(() => AriaReportDocumentBuilder.Build(null));
            TestAssert.True(missing.InnerException == null);
        }

        public static void DotSegmentsCannotBecomeRelativePatientOrProviderTargets()
        {
            Reject(() => Request(patientReference: "Patient/."));
            Reject(() => Request(patientReference: "Patient/.."));
            Reject(() => Request(organizationReference: "Organization/."));
            Reject(() => Request(organizationReference: "Organization/.."));
        }

        public static void UnicodeLineBreaksAndDriveRelativePathsAreRejected()
        {
            foreach (string separator in new[] { "\u2028", "\u2029" })
            {
                Reject(() => Request(description: "Synthetic" + separator + "description"));
                Reject(() => Request(title: "Synthetic" + separator + "review.pdf"));
                Reject(() => Request(documentTypeDisplay: "Synthetic" + separator + "report"));
            }
            Reject(() => Request(description: "C:private.pdf"));
            Reject(() => Request(description: "Prepared from C:private.pdf"));
            int punctuationPathsRejected = 0;
            foreach (string path in new[] { "(C:private.pdf)", "Source=C:private.pdf", "\"C:private.pdf\"" })
            {
                try { Request(description: path); }
                catch (ArgumentException) { punctuationPathsRejected++; }
            }
            TestAssert.Equal(3, punctuationPathsRejected,
                "Drive-relative paths after punctuation must all be rejected.");
            TestAssert.Equal("Review: synthetic", Request(description: "Review: synthetic").Description);
            TestAssert.Equal("Plan review: complete", (string)AriaReportDocumentBuilder.Build(
                Request(description: "Plan review: complete"))["description"]);
        }

        private static byte[] Pdf()
        {
            return Encoding.ASCII.GetBytes("%PDF-1.4\n% SYNTHETIC FORMAT FIXTURE ONLY\n%%EOF\n");
        }

        private static void Reject(Action action)
        {
            var error = TestAssert.Throws<ArgumentException>(action);
            TestAssert.True(error.InnerException == null);
        }

        private static AriaReportUploadRequest Request(
            string reportPatientId = "SYNTHETIC-001", string resolvedPatientId = "SYNTHETIC-001",
            string patientReference = "Patient/synthetic-fhir-1", string activePlanKey = "synthetic-course|synthetic-plan",
            string organizationReference = "Organization/synthetic-provider-1",
            string documentTypeSystem = "urn:oid:1.2.3.4", string documentTypeCode = "synthetic-code",
            string documentTypeDisplay = "Synthetic report", byte[] pdfBytes = null,
            string title = "ClearPlan review.pdf", DateTimeOffset? createdUtc = null, string description = null,
            string categoryCode = "PDF", string categoryDisplay = "PDF")
        {
            return new AriaReportUploadRequest(reportPatientId, resolvedPatientId, patientReference, activePlanKey,
                organizationReference, documentTypeSystem, documentTypeCode, documentTypeDisplay,
                pdfBytes ?? Pdf(), title, createdUtc ?? new DateTimeOffset(2026, 7, 27, 12, 15, 16, TimeSpan.Zero),
                description, categoryCode, categoryDisplay);
        }
    }
}
