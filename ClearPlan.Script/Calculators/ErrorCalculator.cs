using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan
{
    public class ErrorCalculator
    {
        public Patient Patient { get; set; }

        public void AddNewRow(string description, string status, string severity, List<ErrorViewModel> errorGrid)
        {
            errorGrid.Add(new ErrorViewModel
            {
                Description = description,
                Status = status,
                Severity = severity
            });
        }

        public void AddNewRow2(string refPointId, double prescription, double session, double fractions, double d50, double d95, double d2, double d98, List<RefViewModel> refGrid)
        {
            refGrid.Add(new RefViewModel
            {
                RefPointId = refPointId,
                Prescription = prescription,
                Session = session,
                Fractions = fractions,
                D50 = d50,
                D95 = d95,
                D2 = d2,
                D98 = d98
            });
        }

        public List<ErrorViewModel> Calculate(PlanningItem planningItem, Patient patient)
        {
            var errorGrid = new List<ErrorViewModel>();

            try
            {
                if (planningItem is PlanSetup planSetup)
                {
                    errorGrid.AddRange(GetPlanSetupErrors(planSetup, patient));
                }
                else if (planningItem is PlanSum planSum)
                {
                    errorGrid.AddRange(GetPlanSumRxErrors(planSum));
                    foreach (PlanSetup planSetupInSum in planSum.PlanSetups)
                    {
                        errorGrid.AddRange(GetPlanSetupErrors(planSetupInSum, patient));
                    }
                }
            }
            catch (Exception ex)
            {
                AddNewRow("ClearPlan could not finish all starter plan checks: " + ex.Message, "0 - Information", "SYSTEM-001", errorGrid);
            }

            return errorGrid
                .OrderBy(row => row.Status)
                .ThenBy(row => row.Severity)
                .ToList();
        }

        public List<RefViewModel> Calculate2(PlanningItem planningItem, Patient patient)
        {
            var refGrid = new List<RefViewModel>();

            try
            {
                if (planningItem is PlanSetup planSetup)
                {
                    refGrid.AddRange(GetRefData(planSetup, patient));
                }
                else if (planningItem is PlanSum planSum)
                {
                    foreach (PlanSetup planSetupInSum in planSum.PlanSetups)
                    {
                        refGrid.AddRange(GetRefData(planSetupInSum, patient));
                    }
                }
            }
            catch
            {
            }

            return refGrid;
        }

        public List<ErrorViewModel> GetPlanSumRxErrors(PlanSum planSum)
        {
            var errorGrid = new List<ErrorViewModel>();
            var treatmentUnits = new List<string>();
            var xIso = new List<double>();
            var yIso = new List<double>();
            var zIso = new List<double>();
            var planNames = new List<string>();

            foreach (PlanSetup planSetup in planSum.PlanSetups.OrderBy(x => x.CreationDateTime))
            {
                Beam firstTreatmentBeam = planSetup.Beams.FirstOrDefault(beam => !beam.IsSetupField);
                if (firstTreatmentBeam == null)
                {
                    continue;
                }

                treatmentUnits.Add(firstTreatmentBeam.TreatmentUnit.Id);
                xIso.Add(Math.Round(firstTreatmentBeam.IsocenterPosition.x, 2));
                yIso.Add(Math.Round(firstTreatmentBeam.IsocenterPosition.y, 2));
                zIso.Add(Math.Round(firstTreatmentBeam.IsocenterPosition.z, 2));
                planNames.Add(planSetup.UseGating ? planSetup.Id + " (gated)" : planSetup.Id);
            }

            if (!planNames.Any())
            {
                AddNewRow("PlanSum does not contain a treatment beam that can be evaluated.", "1 - Deviation", "SUM-001", errorGrid);
                return errorGrid;
            }

            if (treatmentUnits.Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1)
            {
                AddNewRow("Different treatment units are used inside the PlanSum.", "1 - Deviation", "SUM-002", errorGrid);
            }

            bool consistentIso = xIso.Distinct().Count() == 1 && yIso.Distinct().Count() == 1 && zIso.Distinct().Count() == 1;
            if (!consistentIso)
            {
                AddNewRow("The isocenter is not consistent across all plans in the PlanSum.", "2 - Variation", "SUM-003", errorGrid);
            }

            AddNewRow("PlanSum contains: " + string.Join(" + ", planNames), "0 - Information", "SUM-004", errorGrid);

            if (!errorGrid.Any(row => row.Status != "0 - Information"))
            {
                AddNewRow("No starter PlanSum issues were detected.", "3 - OK", "SUM-000", errorGrid);
            }

            return errorGrid;
        }

        public List<RefViewModel> GetRefData(PlanSetup planSetup, Patient patient)
        {
            var refGrid = new List<RefViewModel>();
            double totalRxDose = planSetup.TotalDose.Dose;

            foreach (var referencePoint in planSetup.ReferencePoints)
            {
                bool isPrimary = planSetup.PrimaryReferencePoint != null && planSetup.PrimaryReferencePoint.Id == referencePoint.Id;
                Structure matchedStructure = planSetup.StructureSet.Structures
                    .FirstOrDefault(structure => structure.Id.ToLower().Replace(" ", "").StartsWith(referencePoint.Id.ToLower().Replace(" ", "")));

                if (matchedStructure != null)
                {
                    AddReferenceRow(planSetup, referencePoint, matchedStructure, totalRxDose, isPrimary, refGrid);
                }
                else
                {
                    AddNewRow2(
                        FormatReferenceLabel(referencePoint.Id, planSetup.Id, isPrimary),
                        referencePoint.TotalDoseLimit.Dose,
                        referencePoint.SessionDoseLimit.Dose,
                        planSetup.NumberOfFractions.HasValue ? planSetup.NumberOfFractions.Value : 0,
                        double.NaN,
                        double.NaN,
                        double.NaN,
                        double.NaN,
                        refGrid);
                }
            }

            if (!refGrid.Any() && !string.IsNullOrWhiteSpace(planSetup.TargetVolumeID))
            {
                Structure target = planSetup.StructureSet.Structures.FirstOrDefault(structure => structure.Id == planSetup.TargetVolumeID);
                if (target != null)
                {
                    AddNewRow2(
                        string.Format("[Target] {0} ({1})", planSetup.TargetVolumeID, planSetup.Id),
                        planSetup.TotalDose.Dose,
                        planSetup.DosePerFraction.Dose,
                        planSetup.NumberOfFractions.HasValue ? planSetup.NumberOfFractions.Value : 0,
                        SafeGetDoseAtVolume(planSetup, target, 50, VolumePresentation.Relative, DoseValuePresentation.Absolute),
                        SafeScaleRelativeDose(planSetup, target, 95, totalRxDose, planSetup.TotalDose.Dose),
                        SafeScaleRelativeDose(planSetup, target, 2, totalRxDose, planSetup.TotalDose.Dose),
                        SafeScaleRelativeDose(planSetup, target, 98, totalRxDose, planSetup.TotalDose.Dose),
                        refGrid);
                }
            }

            return refGrid;
        }

        public List<ErrorViewModel> GetPlanSetupErrors(PlanSetup planSetup, Patient patient)
        {
            var errorGrid = new List<ErrorViewModel>();

            if (!planSetup.IsDoseValid())
            {
                AddNewRow("Plan '" + planSetup.Id + "' does not have a valid dose.", "1 - Deviation", "PLAN-001", errorGrid);
                return errorGrid;
            }

            List<Beam> treatmentBeams = planSetup.Beams.Where(beam => !beam.IsSetupField).ToList();
            if (!treatmentBeams.Any())
            {
                AddNewRow("Plan '" + planSetup.Id + "' has no treatment beam.", "1 - Deviation", "PLAN-002", errorGrid);
                return errorGrid;
            }

            AddApprovalCheck(planSetup, errorGrid);
            AddTargetChecks(planSetup, errorGrid);
            AddReferencePointChecks(planSetup, errorGrid);
            AddTreatmentUnitCheck(planSetup, treatmentBeams, errorGrid);
            AddDoseGridChecks(planSetup, errorGrid);
            AddImageResolutionCheck(planSetup, errorGrid);

            if (!errorGrid.Any())
            {
                AddNewRow("No starter plan-check issues were detected for plan '" + planSetup.Id + "'.", "3 - OK", "PLAN-000", errorGrid);
            }

            return errorGrid;
        }

        private void AddApprovalCheck(PlanSetup planSetup, List<ErrorViewModel> errorGrid)
        {
            if (planSetup.ApprovalStatus == PlanSetupApprovalStatus.UnApproved)
            {
                AddNewRow("Plan '" + planSetup.Id + "' is still unapproved.", "2 - Variation", "PLAN-010", errorGrid);
            }
        }

        private void AddTargetChecks(PlanSetup planSetup, List<ErrorViewModel> errorGrid)
        {
            if (string.IsNullOrWhiteSpace(planSetup.TargetVolumeID))
            {
                AddNewRow("Plan '" + planSetup.Id + "' has no target volume assigned.", "2 - Variation", "PLAN-020", errorGrid);
                return;
            }

            Structure targetStructure = planSetup.StructureSet.Structures.FirstOrDefault(structure => structure.Id == planSetup.TargetVolumeID);
            if (targetStructure == null || targetStructure.IsEmpty)
            {
                AddNewRow("Target volume '" + planSetup.TargetVolumeID + "' could not be matched to a non-empty structure in plan '" + planSetup.Id + "'.", "1 - Deviation", "PLAN-021", errorGrid);
            }
        }

        private void AddReferencePointChecks(PlanSetup planSetup, List<ErrorViewModel> errorGrid)
        {
            if (planSetup.PrimaryReferencePoint == null)
            {
                AddNewRow("Plan '" + planSetup.Id + "' has no primary reference point.", "2 - Variation", "PLAN-030", errorGrid);
                return;
            }

            double totalDoseDifference = Math.Abs(planSetup.PrimaryReferencePoint.TotalDoseLimit.Dose - planSetup.TotalDose.Dose);
            if (totalDoseDifference > 0.05)
            {
                AddNewRow(
                    string.Format(
                        "Primary reference point total dose ({0:0.###} Gy) does not match plan total dose ({1:0.###} Gy) in plan '{2}'.",
                        planSetup.PrimaryReferencePoint.TotalDoseLimit.Dose,
                        planSetup.TotalDose.Dose,
                        planSetup.Id),
                    "1 - Deviation",
                    "PLAN-031",
                    errorGrid);
            }

            double dosePerFractionDifference = Math.Abs(planSetup.PrimaryReferencePoint.SessionDoseLimit.Dose - planSetup.DosePerFraction.Dose);
            if (dosePerFractionDifference > 0.05)
            {
                AddNewRow(
                    string.Format(
                        "Primary reference point dose per fraction ({0:0.###} Gy) does not match plan dose per fraction ({1:0.###} Gy) in plan '{2}'.",
                        planSetup.PrimaryReferencePoint.SessionDoseLimit.Dose,
                        planSetup.DosePerFraction.Dose,
                        planSetup.Id),
                    "1 - Deviation",
                    "PLAN-032",
                    errorGrid);
            }
        }

        private void AddTreatmentUnitCheck(PlanSetup planSetup, List<Beam> treatmentBeams, List<ErrorViewModel> errorGrid)
        {
            if (treatmentBeams.Select(beam => beam.TreatmentUnit.Id).Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1)
            {
                AddNewRow("Plan '" + planSetup.Id + "' uses more than one treatment unit.", "2 - Variation", "PLAN-040", errorGrid);
            }
        }

        private void AddDoseGridChecks(PlanSetup planSetup, List<ErrorViewModel> errorGrid)
        {
            double doseGridMm;
            if (!TryGetDoseGridMm(planSetup, out doseGridMm))
            {
                return;
            }

            if (planSetup.DosePerFraction.Dose >= 7 && doseGridMm > 2)
            {
                AddNewRow(
                    string.Format("Plan '{0}' uses a {1:0.##} mm dose grid for a fraction dose >= 7 Gy.", planSetup.Id, doseGridMm),
                    "1 - Deviation",
                    "PLAN-050",
                    errorGrid);
            }

            if (planSetup.StructureSet != null && planSetup.StructureSet.Image != null && doseGridMm > planSetup.StructureSet.Image.ZRes)
            {
                AddNewRow(
                    string.Format("Plan '{0}' uses a dose-grid z-resolution ({1:0.##} mm) larger than the image z-resolution ({2:0.##} mm).", planSetup.Id, doseGridMm, planSetup.StructureSet.Image.ZRes),
                    "2 - Variation",
                    "PLAN-051",
                    errorGrid);
            }
        }

        private void AddImageResolutionCheck(PlanSetup planSetup, List<ErrorViewModel> errorGrid)
        {
            if (planSetup.StructureSet == null || planSetup.StructureSet.Image == null)
            {
                return;
            }

            if (planSetup.DosePerFraction.Dose > 5 && planSetup.StructureSet.Image.ZRes > 2)
            {
                AddNewRow(
                    string.Format("Plan '{0}' has a fraction dose > 5 Gy with an image z-resolution of {1:0.##} mm.", planSetup.Id, planSetup.StructureSet.Image.ZRes),
                    "2 - Variation",
                    "PLAN-060",
                    errorGrid);
            }
        }

        private static bool TryGetDoseGridMm(PlanSetup planSetup, out double doseGridMm)
        {
            doseGridMm = 0;
            var gridSizeEntry = planSetup.PhotonCalculationOptions.FirstOrDefault(option => option.Key == "CalculationGridSizeInCM");
            if (gridSizeEntry.Key == null)
            {
                return false;
            }

            double doseGridCm;
            if (!double.TryParse(gridSizeEntry.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out doseGridCm) &&
                !double.TryParse(gridSizeEntry.Value, NumberStyles.Any, CultureInfo.CurrentCulture, out doseGridCm))
            {
                return false;
            }

            doseGridMm = doseGridCm * 10.0;
            return true;
        }

        private void AddReferenceRow(PlanSetup planSetup, ReferencePoint referencePoint, Structure matchedStructure, double totalRxDose, bool isPrimary, List<RefViewModel> refGrid)
        {
            double prescription = referencePoint.TotalDoseLimit.Dose;
            AddNewRow2(
                FormatReferenceLabel(matchedStructure.Id, planSetup.Id, isPrimary),
                prescription,
                referencePoint.SessionDoseLimit.Dose,
                planSetup.NumberOfFractions.HasValue ? planSetup.NumberOfFractions.Value : 0,
                SafeGetDoseAtVolume(planSetup, matchedStructure, 50, VolumePresentation.Relative, DoseValuePresentation.Absolute),
                SafeScaleRelativeDose(planSetup, matchedStructure, 95, totalRxDose, prescription),
                SafeScaleRelativeDose(planSetup, matchedStructure, 2, totalRxDose, prescription),
                SafeScaleRelativeDose(planSetup, matchedStructure, 98, totalRxDose, prescription),
                refGrid);
        }

        private static string FormatReferenceLabel(string label, string planId, bool isPrimary)
        {
            return string.Format("{0}{1} ({2})", isPrimary ? "[PRIMARY] " : string.Empty, label, planId);
        }

        private static double SafeGetDoseAtVolume(PlanSetup planSetup, Structure structure, double volumeValue, VolumePresentation volumePresentation, DoseValuePresentation dosePresentation)
        {
            try
            {
                return planSetup.GetDoseAtVolume(structure, volumeValue, volumePresentation, dosePresentation).Dose;
            }
            catch
            {
                return double.NaN;
            }
        }

        private static double SafeScaleRelativeDose(PlanSetup planSetup, Structure structure, double volumeValue, double totalRxDose, double referenceDose)
        {
            if (referenceDose == 0)
            {
                return double.NaN;
            }

            double relativeDose = SafeGetDoseAtVolume(planSetup, structure, volumeValue, VolumePresentation.Relative, DoseValuePresentation.Relative);
            return double.IsNaN(relativeDose) ? double.NaN : relativeDose * totalRxDose / referenceDose;
        }

        public DateTime? GetFirstDateFromString(string input)
        {
            DateTime dateValue;

            foreach (Match match in Regex.Matches(input, @"[0-9]{4}[0-9]{2}[0-9]{2}"))
            {
                if (DateTime.TryParseExact(match.Value, "yyyyMMdd", null, DateTimeStyles.None, out dateValue))
                {
                    return dateValue;
                }
            }

            return null;
        }
    }
}
