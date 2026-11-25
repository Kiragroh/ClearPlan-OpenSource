using System;
using System.Globalization;
using Newtonsoft.Json.Linq;

namespace ClearPlan.Core.Integration
{
    public static class AriaReportDocumentBuilder
    {
        /// <summary>Creates a fresh preliminary ARIA payload without file, network, or ESAPI access.</summary>
        public static JObject Build(AriaReportUploadRequest request)
        {
            if (request == null) throw new ArgumentException("A prepared ARIA report request is required.");
            string timestamp = request.CreatedUtc.UtcDateTime.ToString(
                "yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture);
            var payload = new JObject
            {
                ["resourceType"] = "DocumentReference",
                ["meta"] = new JObject
                {
                    ["profile"] = new JArray("http://varian.com/fhir/v1/StructureDefinition/DocumentReference")
                },
                ["extension"] = new JArray(new JObject
                {
                    ["url"] = "http://varian.com/fhir/v1/StructureDefinition/documentreference-documentLocation",
                    ["valueString"] = "file-server"
                }),
                ["status"] = "current",
                ["docStatus"] = "preliminary",
                ["subject"] = new JObject { ["reference"] = request.PatientReference },
                ["date"] = request.DocumentDateUtc.UtcDateTime.ToString(
                    "yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture),
                ["type"] = new JObject
                {
                    ["coding"] = new JArray(new JObject
                    {
                        ["system"] = request.DocumentTypeSystem,
                        ["code"] = request.DocumentTypeCode,
                        ["display"] = request.DocumentTypeDisplay
                    })
                },
                ["custodian"] = new JObject { ["reference"] = request.OrganizationReference },
                ["category"] = new JArray(new JObject
                {
                    ["coding"] = new JArray(new JObject
                    {
                        ["system"] = "http://varian.com/fhir/CodeSystem/DocumentReference/documentreference-class",
                        ["code"] = request.CategoryCode,
                        ["display"] = request.CategoryDisplay
                    })
                }),
                ["content"] = new JArray(new JObject
                {
                    ["attachment"] = new JObject
                    {
                        ["contentType"] = "application/pdf",
                        ["data"] = Convert.ToBase64String(request.PdfBytes),
                        ["title"] = request.Title,
                        ["creation"] = timestamp
                    }
                })
            };
            // ARIA assigns id; its profile forbids identifier. The plan binding and SHA256
            // remain local request metadata, not invented FHIR identities or extensions.
            if (request.Description != null) payload["description"] = request.Description;
            if (request.TemplateName != null)
            {
                ((JArray)payload["extension"]).Add(new JObject
                {
                    ["url"] = "http://varian.com/fhir/v1/StructureDefinition/documentreference-templateName",
                    ["valueString"] = request.TemplateName
                });
                payload["author"] = new JArray(new JObject { ["reference"] = request.AuthorReference });
            }
            return payload;
        }
    }
}
