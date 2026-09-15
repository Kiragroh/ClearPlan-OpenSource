using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace ClearPlan.Core.Integration
{
    /// <summary>
    /// Detached binding of caller-resolved patient/plan metadata to one prepared PDF.
    /// The caller must generate the PDF from that patient and recheck the active context
    /// before sending. The signature gate does not verify patient text inside a PDF.
    /// </summary>
    public sealed class AriaReportUploadRequest
    {
        public const int MaximumPdfBytes = 20 * 1024 * 1024;
        private readonly byte[] pdfBytes;

        public AriaReportUploadRequest(string reportPatientId, string resolvedPatientId,
            string patientReference, string activePlanKey, string organizationReference,
            string documentTypeSystem, string documentTypeCode, string documentTypeDisplay,
            byte[] pdfBytes, string title, DateTimeOffset createdUtc,
            string description = null, string categoryCode = "PDF", string categoryDisplay = "PDF")
            : this(reportPatientId, resolvedPatientId, patientReference, activePlanKey, organizationReference,
                documentTypeSystem, documentTypeCode, documentTypeDisplay, pdfBytes, title, createdUtc,
                description, categoryCode, categoryDisplay, null)
        {
            // Preserve the original CLR signature for already compiled callers.
        }

        public AriaReportUploadRequest(string reportPatientId, string resolvedPatientId,
            string patientReference, string activePlanKey, string organizationReference,
            string documentTypeSystem, string documentTypeCode, string documentTypeDisplay,
            byte[] pdfBytes, string title, DateTimeOffset createdUtc,
            string description, string categoryCode, string categoryDisplay,
            DateTimeOffset? documentDateUtc = null)
        {
            ValidateText(reportPatientId, 256, "The report patient identity is invalid.", false);
            ValidateText(resolvedPatientId, 256, "The resolved ARIA patient identity is invalid.", false);
            if (!string.Equals(reportPatientId, resolvedPatientId, StringComparison.Ordinal))
                throw new ArgumentException("The report patient does not match the resolved ARIA patient.");

            ValidateReference(patientReference, "Patient");
            ValidateReference(organizationReference, "Organization");
            ValidateText(activePlanKey, 512, "The active plan binding is invalid.", false);
            ValidateSystem(documentTypeSystem);
            ValidateText(documentTypeCode, 256, "The resolved document type code is invalid.", false);
            ValidateText(documentTypeDisplay, 256, "The resolved document type display is invalid.", false);
            if (categoryCode != "PDF" && categoryCode != "TIF")
                throw new ArgumentException("The document category must be a resolved PDF or TIF category.");
            ValidateText(categoryDisplay, 256, "The resolved document category display is invalid.", false);
            if (createdUtc == default(DateTimeOffset))
                throw new ArgumentException("A report creation instant is required.");
            if (documentDateUtc.HasValue && (documentDateUtc.Value == default(DateTimeOffset) || documentDateUtc.Value > createdUtc))
                throw new ArgumentException("The ARIA document date must be valid and no later than report creation.");
            if (pdfBytes == null || pdfBytes.Length < 5 || pdfBytes.Length > MaximumPdfBytes)
                throw new ArgumentException("The report PDF must contain between 5 bytes and 20 MiB.");

            // Validate the retained copy so subsequent caller mutation cannot change the gate or hash.
            this.pdfBytes = (byte[])pdfBytes.Clone();
            if (this.pdfBytes[0] != '%' || this.pdfBytes[1] != 'P' || this.pdfBytes[2] != 'D' ||
                this.pdfBytes[3] != 'F' || this.pdfBytes[4] != '-')
                throw new ArgumentException("The report data does not have a PDF signature.");

            ReportPatientId = reportPatientId;
            ResolvedPatientId = resolvedPatientId;
            PatientReference = patientReference;
            ActivePlanKey = activePlanKey;
            OrganizationReference = organizationReference;
            DocumentTypeSystem = documentTypeSystem;
            DocumentTypeCode = documentTypeCode;
            DocumentTypeDisplay = documentTypeDisplay;
            Title = NormalizeTitle(title);
            CreatedUtc = createdUtc.ToUniversalTime();
            DocumentDateUtc = (documentDateUtc ?? createdUtc).ToUniversalTime();
            Description = NormalizeDescription(description);
            CategoryCode = categoryCode;
            CategoryDisplay = categoryDisplay;
            try
            {
                using (var sha = SHA256.Create())
                    PdfSha256 = BitConverter.ToString(sha.ComputeHash(this.pdfBytes)).Replace("-", "").ToLowerInvariant();
            }
            catch (CryptographicException)
            {
                throw new ArgumentException("The report fingerprint could not be created.");
            }
        }

        public string ReportPatientId { get; private set; }
        public string ResolvedPatientId { get; private set; }
        public string PatientReference { get; private set; }
        public string ActivePlanKey { get; private set; }
        public string OrganizationReference { get; private set; }
        public string DocumentTypeSystem { get; private set; }
        public string DocumentTypeCode { get; private set; }
        public string DocumentTypeDisplay { get; private set; }
        public byte[] PdfBytes { get { return (byte[])pdfBytes.Clone(); } }
        /// <summary>Lowercase SHA256 of the retained bytes for local confirmation/provenance, not FHIR attachment.hash.</summary>
        public string PdfSha256 { get; private set; }
        public int PdfLength { get { return pdfBytes.Length; } }
        public string Title { get; private set; }
        public DateTimeOffset CreatedUtc { get; private set; }
        /// <summary>Explicit ARIA document date, which may precede but never replaces report creation.</summary>
        public DateTimeOffset DocumentDateUtc { get; private set; }
        public string Description { get; private set; }
        public string CategoryCode { get; private set; }
        public string CategoryDisplay { get; private set; }
        public string TemplateName { get; private set; }
        public string AuthorReference { get; private set; }
        public string InteractiveUserId { get; private set; }

        /// <summary>Binds exact native plan/user metadata without mutating the prepared PDF or legacy request.</summary>
        public AriaReportUploadRequest WithClinicalMetadata(string activePlanName, string authorReference, string interactiveUserId)
        {
            // TemplateName is ARIA metadata, not a filename: retain spaces, Unicode and punctuation exactly.
            ValidateText(activePlanName, 252, "The active plan name is missing or invalid for ARIA TemplateName.", false);
            ValidateReference(authorReference, "Practitioner");
            ValidateText(interactiveUserId, 256, "The interactive ESAPI user identity is missing or invalid.", false);
            var bound = new AriaReportUploadRequest(ReportPatientId, ResolvedPatientId, PatientReference, ActivePlanKey,
                OrganizationReference, DocumentTypeSystem, DocumentTypeCode, DocumentTypeDisplay, pdfBytes, Title,
                CreatedUtc, Description, CategoryCode, CategoryDisplay, DocumentDateUtc);
            bound.TemplateName = "PQM_" + activePlanName;
            bound.AuthorReference = authorReference;
            bound.InteractiveUserId = interactiveUserId;
            return bound;
        }

        private static void ValidateReference(string value, string resourceType)
        {
            if (value == null || value == resourceType + "/." || value == resourceType + "/.." ||
                !Regex.IsMatch(value, "\\A" + resourceType + "/[A-Za-z0-9.-]{1,64}\\z",
                    RegexOptions.CultureInvariant))
                throw new ArgumentException("A resolved relative FHIR resource reference is invalid.");
        }

        private static void ValidateSystem(string value)
        {
            const string error = "The resolved document type system is invalid.";
            ValidateText(value, 512, error, false);
            foreach (char character in value)
                if (char.IsWhiteSpace(character) || character == '\\') throw new ArgumentException(error);
            Uri uri;
            if (!Uri.TryCreate(value, UriKind.Absolute, out uri) ||
                (uri.Scheme != "http" && uri.Scheme != "https" && uri.Scheme != "urn") ||
                !string.IsNullOrEmpty(uri.UserInfo) || !string.IsNullOrEmpty(uri.Query) || !string.IsNullOrEmpty(uri.Fragment) ||
                (uri.Scheme == "urn" && value.Length <= 4))
                throw new ArgumentException(error);
        }

        private static string NormalizeTitle(string value)
        {
            const string error = "The report title must be a concise PDF name without a path.";
            ValidateText(value, 200, error, true);
            if (value.IndexOfAny(new[] { '/', '\\', ':', '"', '|', '?', '*' }) >= 0)
                throw new ArgumentException(error);
            string normalized = Regex.Replace(value.Trim(), " +", " ");
            if (normalized.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                normalized = normalized.Substring(0, normalized.Length - 4).TrimEnd();
            if (string.IsNullOrWhiteSpace(normalized) || normalized[0] == '.' ||
                normalized.EndsWith(".", StringComparison.Ordinal) || normalized.Length + 4 > 200)
                throw new ArgumentException(error);
            return normalized + ".pdf";
        }

        private static string NormalizeDescription(string value)
        {
            if (value == null || value.Length == 0) return null;
            const string error = "The report description must be concise text without paths or control characters.";
            if (value.Length > 512 || HasUnsafeCharacters(value) || value.IndexOfAny(new[] { '/', '\\' }) >= 0 ||
                Regex.IsMatch(value, @"(?<![A-Za-z0-9])[A-Za-z]:\S", RegexOptions.CultureInvariant))
                throw new ArgumentException(error);
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }

        private static void ValidateText(string value, int maximumLength, string error, bool allowEdgeSpaces)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Length > maximumLength || HasUnsafeCharacters(value) ||
                (!allowEdgeSpaces && value != value.Trim()))
                throw new ArgumentException(error);
        }

        private static bool HasUnsafeCharacters(string value)
        {
            foreach (char character in value)
            {
                UnicodeCategory category = char.GetUnicodeCategory(character);
                if (char.IsControl(character) || category == UnicodeCategory.Format ||
                    category == UnicodeCategory.LineSeparator || category == UnicodeCategory.ParagraphSeparator ||
                    category == UnicodeCategory.Surrogate || character == '<' || character == '>') return true;
            }
            return false;
        }
    }
}
