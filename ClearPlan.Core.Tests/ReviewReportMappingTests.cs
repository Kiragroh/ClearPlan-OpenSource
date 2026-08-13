using System;
using System.IO;
using System.Linq;
using System.Reflection;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;
using ClearPlan.Reporting;
using ClearPlan.Reporting.MigraDoc;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.IO;

namespace ClearPlan.Core.Tests
{
    internal static class ReviewReportMappingTests
    {
        private const string SyntheticWatermark =
            "SYNTHETIC DEMONSTRATION — NOT FOR CLINICAL USE";

        public static void PreservesEveryReviewSectionAndUnit()
        {
            ReviewSnapshot fieldSnapshot =
                SyntheticScenarioFactory.Create("field-and-mapping");
            ReviewSnapshot metadataSnapshot =
                SyntheticScenarioFactory.Create("metadata-plancheck");
            var mapper = new ReviewSnapshotReportMapper();

            ReviewReportDocument report = mapper.Map(fieldSnapshot);
            ReviewReportDocument metadataReport = mapper.Map(metadataSnapshot);

            TestAssert.Equal(fieldSnapshot.Plans.Count, report.Plans.Count);
            TestAssert.Equal(
                fieldSnapshot.ActivePlanKey,
                report.ActivePlanKey);
            TestAssert.Equal(
                fieldSnapshot.Plans[0].DosePerFractionGy,
                report.Plans[0].DosePerFractionGy);
            TestAssert.Equal(
                fieldSnapshot.Plans[0].TotalDoseGy,
                report.Plans[0].TotalDoseGy);
            TestAssert.Equal(
                fieldSnapshot.Plans[0].FractionCount,
                report.Plans[0].FractionCount);

            ReviewPqmRow sourcePqm = fieldSnapshot.PqmRows[0];
            ReviewReportPqmRow reportPqm = report.PqmRows[0];
            TestAssert.Equal(sourcePqm.StableId, reportPqm.StableId);
            TestAssert.Equal(sourcePqm.TemplateCode, reportPqm.TemplateCode);
            TestAssert.Equal(
                sourcePqm.TemplateStructure,
                reportPqm.TemplateStructure);
            TestAssert.Equal(
                sourcePqm.ResolvedStructureId,
                reportPqm.ResolvedStructureId);
            TestAssert.Equal(sourcePqm.Objective, reportPqm.Objective);
            TestAssert.Equal(sourcePqm.Comparator, reportPqm.Comparator);
            TestAssert.Equal(sourcePqm.Goal, reportPqm.Goal);
            TestAssert.Equal(sourcePqm.Variation, reportPqm.Variation);
            TestAssert.Equal(
                sourcePqm.AchievedValue,
                reportPqm.AchievedValue);
            TestAssert.Equal(sourcePqm.Unit, reportPqm.Unit);
            TestAssert.Equal(sourcePqm.Status, reportPqm.Status);
            TestAssert.Equal(sourcePqm.Severity, reportPqm.Severity);
            TestAssert.Equal(
                sourcePqm.StructureOptions.Count,
                reportPqm.StructureOptions.Count);

            ReviewCheckRow sourceCheck = metadataSnapshot.PlanCheckRows
                .Single(row => row.CheckCode == "appointment-present");
            ReviewReportCheckRow reportCheck = metadataReport.PlanCheckRows
                .Single(row => row.CheckCode == "appointment-present");
            TestAssert.Equal(sourceCheck.Category, reportCheck.Category);
            TestAssert.Equal(sourceCheck.ObservedValue, reportCheck.ObservedValue);
            TestAssert.Equal(sourceCheck.ExpectedValue, reportCheck.ExpectedValue);
            TestAssert.Equal(sourceCheck.Unit, reportCheck.Unit);
            TestAssert.Equal(sourceCheck.Status, reportCheck.Status);
            TestAssert.Equal(sourceCheck.Severity, reportCheck.Severity);
            TestAssert.Equal(sourceCheck.Message, reportCheck.Message);

            ReviewFieldRow sourceField = fieldSnapshot.FieldRows[0];
            ReviewReportFieldRow reportField = report.FieldRows[0];
            TestAssert.Equal(sourceField.CurrentId, reportField.CurrentId);
            TestAssert.Equal(sourceField.ExpectedId, reportField.ExpectedId);
            TestAssert.Equal(sourceField.CurrentName, reportField.CurrentName);
            TestAssert.Equal(
                sourceField.SuggestedName,
                reportField.SuggestedName);
            TestAssert.Equal(sourceField.IdStatus, reportField.IdStatus);
            TestAssert.Equal(sourceField.NameStatus, reportField.NameStatus);

            ReviewDvhSeries sourceDvh = fieldSnapshot.DvhSeries[0];
            ReviewReportDvhSeries reportDvh = report.DvhSeries[0];
            TestAssert.Equal(ReviewUnitCodes.Gray, reportDvh.DoseUnit);
            TestAssert.Equal(ReviewUnitCodes.Percent, reportDvh.VolumeUnit);
            TestAssert.Equal(sourceDvh.StructureId, reportDvh.StructureId);
            TestAssert.Equal(sourceDvh.VolumeCc, reportDvh.VolumeCc);
            TestAssert.Equal(sourceDvh.Points.Count, reportDvh.Points.Count);
            TestAssert.Equal(
                sourceDvh.Points[10].DoseGy,
                reportDvh.Points[10].DoseGy);
            TestAssert.Equal(
                sourceDvh.Points[10].VolumePercent,
                reportDvh.Points[10].VolumePercent);

            TestAssert.Equal(
                fieldSnapshot.StructureMappings.Count,
                report.StructureMappings.Count);
            TestAssert.Equal(
                fieldSnapshot.Sources.Count,
                report.Sources.Count);
        }

        public static void ForcesTheSyntheticSafetyMark()
        {
            ReviewSnapshot snapshot =
                SyntheticScenarioFactory.Create("baseline-pass");
            snapshot.Report.Watermark = string.Empty;
            snapshot.Report.ModeLabel = string.Empty;

            ReviewReportDocument report =
                new ReviewSnapshotReportMapper().Map(snapshot);

            TestAssert.True(report.Synthetic);
            TestAssert.Equal(SyntheticWatermark, report.Watermark);
            TestAssert.Equal(SyntheticWatermark, report.ModeLabel);
            TestAssert.Equal(snapshot.ScenarioId, report.ScenarioId);
            TestAssert.Equal(snapshot.ProvenanceText, report.ProvenanceText);
            TestAssert.Equal(snapshot.Report.Notes.Count, report.Notes.Count);
        }

        public static void WritesASyntheticPdf()
        {
            ReviewSnapshot snapshot =
                SyntheticScenarioFactory.Create("metadata-plancheck");
            ReviewReportDocument report =
                new ReviewSnapshotReportMapper().Map(snapshot);
            string outputPath = Path.Combine(
                Path.GetTempPath(),
                "clearplan-report-" + Guid.NewGuid().ToString("N") + ".pdf");

            try
            {
                new ReportPdf().Export(outputPath, report);

                TestAssert.True(File.Exists(outputPath));
                byte[] bytes = File.ReadAllBytes(outputPath);
                TestAssert.True(bytes.Length > 1000);
                TestAssert.Equal((byte)'%', bytes[0]);
                TestAssert.Equal((byte)'P', bytes[1]);
                TestAssert.Equal((byte)'D', bytes[2]);
                TestAssert.Equal((byte)'F', bytes[3]);
            }
            finally
            {
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }
            }
        }

        public static void PopulatesMappedPdfTableCells()
        {
            ReviewSnapshot snapshot =
                SyntheticScenarioFactory.Create("metadata-plancheck");
            ReviewReportDocument report =
                new ReviewSnapshotReportMapper().Map(snapshot);
            MethodInfo createDocument = typeof(ReportPdf).GetMethod(
                "CreateReviewReport",
                BindingFlags.Instance | BindingFlags.NonPublic);
            TestAssert.NotNull(createDocument);

            var document = (Document)createDocument.Invoke(
                new ReportPdf(),
                new object[] { report });
            string ddl = DdlWriter.WriteToString(document);

            TestAssert.True(
                ddl.Contains("Synthetic review plan"),
                "The plan row must contain its mapped display label.");
            TestAssert.True(
                ddl.Contains("appointment-present"),
                "The PlanCheck row must contain its stable check code.");
            TestAssert.True(
                ddl.Contains("D95"),
                "The PQM row must contain its mapped objective.");
            TestAssert.True(
                ddl.Contains("SYN01"),
                "The field row must contain its mapped field ID.");
        }
    }
}
