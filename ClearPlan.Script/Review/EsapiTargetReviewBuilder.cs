using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows;
using ClearPlan.Core.Review;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Review
{
    public sealed class TargetReviewCapture
    {
        public List<TargetReviewStructure> Structures { get; set; } = new List<TargetReviewStructure>();
        public List<string> RequiredStructureIds { get; set; } = new List<string>();
        public string SelectionMessage { get; set; }
    }

    /// <summary>Reads native target contours and exact prescription assignments on the ESAPI owner.</summary>
    public sealed class EsapiTargetReviewBuilder
    {
        private const int MaximumImagePlanes = 4096;
        private const int MaximumContourPoints = 2000000;
        private static readonly TimeSpan CaptureLimit = TimeSpan.FromSeconds(15);

        public TargetReviewCapture BuildSelection(PlanningItem planningItem)
        {
            VerifyOwner();
            var result = new TargetReviewCapture();
            if (planningItem == null || planningItem.StructureSet == null)
            {
                result.SelectionMessage = "Target selection not evaluated: native structure set unavailable.";
                return result;
            }
            var native = new Dictionary<string, Structure>(StringComparer.Ordinal);
            foreach (Structure structure in planningItem.StructureSet.Structures)
            {
                if (structure == null) continue;
                string kind = DvhSelectionPolicy.ClassifyTarget(structure.Id, structure.DicomType);
                if (kind.Length == 0) continue;
                var detached = new TargetReviewStructure
                {
                    StructureId = structure.Id, DicomType = structure.DicomType,
                    IsEmpty = structure.IsEmpty || !structure.HasSegment,
                    VolumeCc = FinitePositive(structure.Volume),
                    ContainmentReason = kind == "PTV" ? "PTV retained independently of OAR/constraint selection."
                        : "Native contour-plane containment not established."
                };
                result.Structures.Add(detached);
                native[structure.Id] = structure;
            }
            var ptvs = result.Structures.Where(item => !item.IsEmpty &&
                DvhSelectionPolicy.ClassifyTarget(item.StructureId, item.DicomType) == "PTV").ToList();
            var inner = result.Structures.Where(item => !item.IsEmpty &&
                DvhSelectionPolicy.ClassifyTarget(item.StructureId, item.DicomType) != "PTV").ToList();
            var image = planningItem.StructureSet.Image;
            var watch = Stopwatch.StartNew();
            int pointCount = 0;
            // Each polygon is read once. This cache contains detached points only.
            var contours = new Dictionary<string, List<Point[]>>();
            if (image != null && image.ZSize > 0 && image.ZSize <= MaximumImagePlanes && ptvs.Count > 0)
            {
                foreach (TargetReviewStructure candidate in inner.OrderByDescending(item => item.VolumeCc ?? double.MaxValue))
                {
                    bool unknown = false;
                    foreach (TargetReviewStructure ptv in ptvs)
                    {
                        try
                        {
                            bool contained = true;
                            bool hasContour = false;
                            for (int plane = 0; plane < image.ZSize; plane++)
                            {
                                if (watch.Elapsed > CaptureLimit) throw new TimeoutException();
                                List<Point[]> inside = ReadPlane(native[candidate.StructureId], image, plane, contours, ref pointCount);
                                if (inside.Count == 0) continue;
                                hasContour = true;
                                List<Point[]> outside = ReadPlane(native[ptv.StructureId], image, plane, contours, ref pointCount);
                                if (!TargetContourContainment.IsContained(inside, outside))
                                {
                                    contained = false;
                                    break;
                                }
                            }
                            if (contained && hasContour)
                            {
                                candidate.FullyContainedInPtv = true;
                                candidate.ContainmentReason = "Fully contained on every native CT contour plane in " + ptv.StructureId +
                                    "; complete polygons including holes, area tolerance 0.000001 mm2. Continuous inter-plane volume was not tested.";
                                break;
                            }
                            if (!hasContour) unknown = true;
                        }
                        catch (Exception)
                        {
                            unknown = true;
                            candidate.ContainmentReason = "Native contour-plane containment unavailable or bounded capture limit reached; no geometric match inferred.";
                        }
                    }
                    if (candidate.FullyContainedInPtv != true)
                    {
                        candidate.FullyContainedInPtv = unknown ? (bool?)null : false;
                        if (!unknown) candidate.ContainmentReason = "Not fully contained in any PTV on native CT contour planes.";
                    }
                }
            }
            result.RequiredStructureIds = DvhSelectionPolicy.SelectTargetStructureIds(result.Structures);
            int innerCount = result.RequiredStructureIds.Count - ptvs.Count;
            result.SelectionMessage = ptvs.Count + " PTV(s) retained; " + innerCount +
                " largest contained CTV/GTV/ITV selected overall (not one per type/PTV). " +
                "Containment compares complete polygons on every native CT contour plane; it does not test continuous inter-plane volume. " +
                inner.Count(item => !item.FullyContainedInPtv.HasValue) + " inner-target containment(s) unavailable. " +
                "No prescription dose is inferred from containment.";
            return result;
        }

        public List<ReviewPqmRow> BuildGoals(PlanningItem planningItem, TargetReviewCapture capture,
            DefaultReviewRuleConfiguration configuration, out ReviewSourceStatus source)
        {
            VerifyOwner();
            var prescriptions = new List<TargetPrescriptionDose>();
            var rows = new List<ReviewPqmRow>();
            var plan = planningItem as PlanSetup;
            string unavailable = null;
            int? fractions = null;
            bool? valid = null;
            if (plan == null) unavailable = "Target Rx rule requires a single plan, not a plan sum.";
            else
            {
                try { fractions = plan.NumberOfFractions; } catch (Exception) { }
                try { valid = plan.Dose != null && plan.IsDoseValid; } catch (Exception) { }
                try
                {
                    var prescription = plan.RTPrescription;
                    if (prescription == null || prescription.Targets == null)
                        unavailable = "Native prescription target metadata is unavailable; no plan-total dose was substituted.";
                    else foreach (var target in prescription.Targets)
                        prescriptions.Add(new TargetPrescriptionDose
                        {
                            TargetId = target.TargetId, TargetType = target.Type.ToString(),
                            DosePerFractionGy = DoseGy(target.DosePerFraction), FractionCount = target.NumberOfFractions
                        });
                }
                catch (Exception)
                {
                    prescriptions.Clear();
                    unavailable = "Native prescription target metadata could not be read; no plan-total dose was substituted.";
                }
            }
            foreach (string id in capture.RequiredStructureIds)
            {
                TargetReviewStructure metadata = capture.Structures.First(item => item.StructureId == id);
                string kind = DvhSelectionPolicy.ClassifyTarget(id, metadata.DicomType);
                foreach (DefaultTargetReviewRule rule in configuration.Rules.Where(item => item.TargetTypes.Contains(kind)))
                {
                    double? achieved = null;
                    double volume;
                    try
                    {
                        Structure structure = planningItem.StructureSet.Structures.Single(item => item.Id == id);
                        if (rule.TryGetVolumePercent(out volume)) achieved = ReadDoseAtVolume(planningItem, structure, volume);
                    }
                    catch (Exception) { }
                    rows.Add(UserTargetCoverageRule.Evaluate(rule, id, prescriptions, fractions, achieved, valid, unavailable));
                }
            }
            source = new ReviewSourceStatus
            {
                StableId = UserTargetCoverageRule.SourceStableId,
                SourceCode = "user-target-coverage", SourceType = "editable-default-rules",
                Status = configuration.Status != ReviewStatusCodes.Available ? configuration.Status :
                    rows.Count == 0 ? ReviewStatusCodes.NotEvaluated : ReviewStatusCodes.Available,
                Optional = false, UsedFallback = false, PathDisplayLabel = UserTargetCoverageRule.SourceLabel,
                Message = configuration.Message + " Category: Default. Independent of stock/RefDB and native assigned goals. " +
                    rows.Count + " row(s); " + rows.Count(row => row.Status == ReviewStatusCodes.NotEvaluated) +
                    " not evaluated. " + capture.SelectionMessage
            };
            return rows;
        }

        public static double? ReadDoseAtVolume(PlanningItem planningItem, Structure structure, double volumePercent)
        {
            VerifyOwner();
            try
            {
                var plan = planningItem as PlanSetup;
                if (plan == null || plan.Dose == null || !plan.IsDoseValid || structure == null || structure.IsEmpty) return null;
                return DoseGy(planningItem.GetDoseAtVolume(structure, volumePercent,
                    VolumePresentation.Relative, DoseValuePresentation.Absolute));
            }
            catch (Exception) { return null; }
        }

        private static List<Point[]> ReadPlane(Structure structure, VMS.TPS.Common.Model.API.Image image, int plane,
            IDictionary<string, List<Point[]>> cache, ref int pointCount)
        {
            string key = structure.Id + "\n" + plane;
            List<Point[]> value;
            if (cache.TryGetValue(key, out value)) return value;
            var native = structure.GetContoursOnImagePlane(plane);
            var x = image.XDirection;
            var y = image.YDirection;
            var origin = image.Origin;
            value = new List<Point[]>();
            foreach (VVector[] contour in native)
            {
                pointCount += contour.Length;
                if (pointCount > MaximumContourPoints) throw new InvalidOperationException("Contour point capture limit.");
                value.Add(contour.Select(point => new Point(
                    (point.x - origin.x) * x.x + (point.y - origin.y) * x.y + (point.z - origin.z) * x.z,
                    (point.x - origin.x) * y.x + (point.y - origin.y) * y.y + (point.z - origin.z) * y.z)).ToArray());
            }
            cache[key] = value;
            return value;
        }

        private static double? DoseGy(DoseValue dose)
        {
            string unit = dose.Unit.ToString();
            if (unit != "Gy" && unit != "cGy") return null;
            double value = ClinicalReviewValueMapper.ConvertDoseToGray(dose.Dose, unit);
            return double.IsNaN(value) || double.IsInfinity(value) || value < 0 ? (double?)null : value;
        }

        private static double? FinitePositive(double value)
        {
            return double.IsNaN(value) || double.IsInfinity(value) || value <= 0 ? (double?)null : value;
        }

        private static void VerifyOwner()
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Native target review requires the ESAPI owner STA.");
            if (System.Windows.Application.Current != null)
                System.Windows.Application.Current.Dispatcher.VerifyAccess();
        }
    }
}
