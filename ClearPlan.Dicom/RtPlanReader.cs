using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClearPlan.Core.PlanAnalysis;
using FellowOakDicom;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace ClearPlan.Dicom
{
    public sealed class ParsedRtPlan
    {
        [JsonIgnore] public string SopClassUid { get; internal set; }
        [JsonIgnore] public string SopInstanceUid { get; internal set; }
        [JsonIgnore] public string FrameOfReferenceUid { get; internal set; }
        [JsonIgnore] public string StudyInstanceUid { get; internal set; }
        public int FractionGroupNumber { get; internal set; }
        public ReviewPlanAnalysis Analysis { get; internal set; }
    }

    public sealed class RtPlanImportException : Exception
    {
        public RtPlanImportException(string code, string safeMessage) : base(safeMessage) { Code = code; }
        public string Code { get; private set; }
    }

    public static class RtPlanReader
    {
        public const string StandardRtPlanSopClass = "1.2.840.10008.5.1.4.1.1.481.5";
        public const string VarianPrivateRtPlanSopClass = "1.2.246.352.70.1.70";
        private static readonly object InitializationLock = new object();
        private static bool _initialized;

        public static ParsedRtPlan Read(string path, int? fractionGroupNumber = null)
        {
            try
            {
                EnsureInitialized();
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path) || new FileInfo(path).Length > 128L * 1024 * 1024)
                    Fail("dicom_read", "The selected RTPLAN is missing, unreadable, or larger than the 128 MiB import limit.");
                DicomFile file = DicomFile.Open(path, FileReadOption.ReadAll);
                if (file == null || file.IsPartial) Fail("dicom_read", "The selected file is not a complete DICOM object.");
                DicomDataset root = file.Dataset;
                string sopClass = Text(root, DicomTag.SOPClassUID);
                if (sopClass != StandardRtPlanSopClass && sopClass != VarianPrivateRtPlanSopClass)
                    Fail("sop_class", "Only standard RT Plan Storage and the documented Varian Private RT Plan Storage class are supported.");
                string sopUid = RequiredUid(root, DicomTag.SOPInstanceUID);
                if (file.FileMetaInfo != null &&
                    ((file.FileMetaInfo.Contains(DicomTag.MediaStorageSOPClassUID) && Text(file.FileMetaInfo, DicomTag.MediaStorageSOPClassUID) != sopClass) ||
                     (file.FileMetaInfo.Contains(DicomTag.MediaStorageSOPInstanceUID) && Text(file.FileMetaInfo, DicomTag.MediaStorageSOPInstanceUID) != sopUid)))
                    Fail("sop_class", "DICOM file metadata and dataset SOP identifiers disagree.");
                if (Text(root, DicomTag.Modality) != "RTPLAN" || Text(root, DicomTag.RTPlanGeometry) != "PATIENT")
                    Fail("plan_geometry", "A patient-coordinate RTPLAN is required.");
                var parsed = new ParsedRtPlan
                {
                    SopClassUid = sopClass, SopInstanceUid = sopUid,
                    StudyInstanceUid = RequiredUid(root, DicomTag.StudyInstanceUID),
                    FrameOfReferenceUid = RequiredUid(root, DicomTag.FrameOfReferenceUID),
                    Analysis = new ReviewPlanAnalysis()
                };

                IList<DicomDataset> groups = Sequence(root, DicomTag.FractionGroupSequence, "fraction_group");
                var groupNumbers = groups.Select(group => Integer(group, DicomTag.FractionGroupNumber, "fraction_group")).ToList();
                if (groupNumbers.Any(number => number <= 0) || groupNumbers.Distinct().Count() != groups.Count)
                    Fail("fraction_group", "Fraction group numbers must be positive and unique.");
                var selected = groups.Where(group => !fractionGroupNumber.HasValue ||
                    Integer(group, DicomTag.FractionGroupNumber, "fraction_group") == fractionGroupNumber.Value).ToList();
                if (selected.Count != 1)
                    Fail("fraction_group", "Select one explicit fraction group when the RTPLAN contains multiple groups; the requested group must exist.");
                DicomDataset fraction = selected[0];
                parsed.FractionGroupNumber = Integer(fraction, DicomTag.FractionGroupNumber, "fraction_group");
                IList<DicomDataset> references = Sequence(fraction, DicomTag.ReferencedBeamSequence, "fraction_group");
                if (Integer(fraction, DicomTag.NumberOfBeams, "fraction_group") != references.Count)
                    Fail("fraction_group", "The fraction group's beam count does not match its references.");
                var allBeams = Sequence(root, DicomTag.BeamSequence, "beam_identity");
                var beamNumbers = allBeams.Select(beam => Integer(beam, DicomTag.BeamNumber, "beam_identity")).ToList();
                if (beamNumbers.Any(number => number <= 0) || beamNumbers.Distinct().Count() != beamNumbers.Count)
                    Fail("beam_identity", "Beam numbers are missing, nonpositive, or duplicated.");
                var used = new HashSet<int>();
                foreach (DicomDataset reference in references)
                {
                    int number = Integer(reference, DicomTag.ReferencedBeamNumber, "beam_identity");
                    if (!used.Add(number) || !beamNumbers.Contains(number))
                        Fail("beam_identity", "A fraction group has duplicate or unknown beam references.");
                    DicomDataset beam = allBeams[beamNumbers.IndexOf(number)];
                    string delivery = Text(beam, DicomTag.TreatmentDeliveryType);
                    if (delivery == "SETUP") continue;
                    if (delivery != "TREATMENT") Fail("beam_type", "Only explicitly identified treatment beams are supported.");
                    double mu = Number(reference, DicomTag.BeamMeterset, "beam_meterset");
                    if (mu < 0 || Text(beam, DicomTag.PrimaryDosimeterUnit) != "MU")
                        Fail("beam_meterset", "Treatment beam metersets must be nonnegative and expressed in MU.");
                    parsed.Analysis.Beams.Add(ReadBeam(beam, number, mu));
                }
                if (parsed.Analysis.Beams.Count == 0) Fail("beam_identity", "The selected fraction group has no supported treatment beams.");
                SetProvenance(parsed.Analysis);
                PlanAnalysisCalculator.Calculate(parsed.Analysis);
                return parsed;
            }
            catch (RtPlanImportException) { throw; }
            catch (Exception)
            {
                // Vendor/parser and filesystem exceptions can contain filenames or DICOM values.
                // Never expose their messages, stack objects, or InnerException in clinical logs.
                throw new RtPlanImportException("dicom_read", "The selected file could not be read as a complete, supported RTPLAN.");
            }
        }

        public static ReviewPlanAnalysis MatchAndApply(ReviewPlanAnalysis active, ParsedRtPlan parsed,
            string expectedSopUid, string expectedFrameUid = null)
        {
            if (parsed == null || string.IsNullOrWhiteSpace(expectedSopUid) ||
                !string.Equals(expectedSopUid, parsed.SopInstanceUid, StringComparison.Ordinal))
                Fail("plan_identity", "The RTPLAN SOP Instance UID does not match the active plan. Nothing was imported.");
            if (expectedFrameUid != null && (string.IsNullOrWhiteSpace(expectedFrameUid) ||
                !string.Equals(expectedFrameUid, parsed.FrameOfReferenceUid, StringComparison.Ordinal)))
                Fail("frame_identity", "The RTPLAN frame of reference does not match the active plan. Nothing was imported.");
            if (active == null || active.Beams == null || parsed.Analysis == null ||
                active.Beams.Count != parsed.Analysis.Beams.Count ||
                active.Beams.Any(beam => !beam.BeamNumber.HasValue) ||
                active.Beams.Select(beam => beam.BeamNumber).Distinct().Count() != active.Beams.Count)
                Fail("beam_match", "The active plan and fraction group must have the same uniquely numbered treatment beams.");
            var sourceByNumber = parsed.Analysis.Beams.ToDictionary(beam => beam.BeamNumber.Value);
            foreach (ReviewBeamAnalysis activeBeam in active.Beams)
            {
                ReviewBeamAnalysis source;
                if (!sourceByNumber.TryGetValue(activeBeam.BeamNumber.Value, out source))
                    Fail("beam_match", "A numbered treatment beam does not match the active plan.");
                ValidateBeamMatch(activeBeam, source);
            }

            // Complete validation precedes any calculation. The input snapshot remains untouched.
            var copy = PlanAnalysisSnapshot.Copy(active);
            var imported = PlanAnalysisSnapshot.Copy(parsed.Analysis).Beams.ToDictionary(beam => beam.BeamNumber.Value);
            for (int beamIndex = 0; beamIndex < active.Beams.Count; beamIndex++)
            {
                var original = active.Beams[beamIndex]; var target = copy.Beams[beamIndex];
                var source = imported[original.BeamNumber.Value];
                target.MlcModel = source.MlcModel;
                target.MetersetMu = source.MetersetMu;
                for (int index = 0; index < target.ControlPoints.Count; index++)
                {
                    var point = target.ControlPoints[index];
                    point.Aperture = source.ControlPoints[index].Aperture;
                    point.PlannedDoseRateMuPerMin = source.ControlPoints[index].PlannedDoseRateMuPerMin;
                    point.GeometryReason = null;
                }
            }
            SetProvenance(copy);
            return PlanAnalysisCalculator.Calculate(copy);
        }

        private static ReviewBeamAnalysis ReadBeam(DicomDataset beam, int number, double mu)
        {
            if (Text(beam, DicomTag.RadiationType) != "PHOTON" ||
                !(Text(beam, DicomTag.BeamType) == "STATIC" || Text(beam, DicomTag.BeamType) == "DYNAMIC"))
                Fail("beam_type", "Only static or dynamic photon treatment beams are supported.");
            if ((beam.Contains(DicomTag.NumberOfBlocks) && Integer(beam, DicomTag.NumberOfBlocks, "beam_type") != 0) ||
                Text(beam, new DicomTag(0x3008, 0x00a3)) == "YES")
                Fail("beam_type", "Block-defined and enhanced beam-limiting-device geometries are not supported by this adapter.");
            var definitions = ReadDevices(beam);
            bool dualLayer = definitions.Any(device => device.Type == "MLCX1") && definitions.Any(device => device.Type == "MLCX2");
            bool fixedBounds = dualLayer && definitions.Any(device => device.Type == "X") && definitions.Any(device => device.Type == "Y");
            if (dualLayer && definitions.Any(device => device.Type == "ASYMX" || device.Type == "ASYMY"))
                Fail("fixed_bounds", "Asymmetric jaw tags alongside the dual-layer extension are not an established supported hardware definition.");
            IList<DicomDataset> cps = Sequence(beam, DicomTag.ControlPointSequence, "control_points");
            if (cps.Count < 2 || cps.Count > 20000 || cps.Count != Integer(beam, DicomTag.NumberOfControlPoints, "control_points"))
                Fail("control_points", "The control point count is invalid or inconsistent.");
            double finalWeight = Number(beam, DicomTag.FinalCumulativeMetersetWeight, "meterset_weights");
            if (finalWeight <= 0) Fail("meterset_weights", "Final cumulative meterset weight must be positive.");
            var result = new ReviewBeamAnalysis { BeamNumber = number, BeamId = "Beam-" + number,
                MetersetMu = mu, MlcModel = "RTPLAN " + string.Join(" + ", definitions.Where(device => device.Axis != null).Select(device => device.Type)) };
            var positions = new Dictionary<string, double[]>(StringComparer.Ordinal);
            double? gantry = null, collimator = null, support = null, rate = null;
            double[] isocenter = null;
            double lastWeight = 0;
            for (int index = 0; index < cps.Count; index++)
            {
                DicomDataset cp = cps[index];
                if (Integer(cp, DicomTag.ControlPointIndex, "control_points") != index)
                    Fail("control_points", "Control point indices must be consecutive and start at zero.");
                ValidateOrientation(cp, cps[0], index);
                gantry = InheritNumber(cp, DicomTag.GantryAngle, gantry);
                collimator = InheritNumber(cp, DicomTag.BeamLimitingDeviceAngle, collimator);
                support = InheritNumber(cp, DicomTag.PatientSupportAngle, support);
                rate = InheritNumber(cp, DicomTag.DoseRateSet, rate);
                if (!gantry.HasValue || !collimator.HasValue || !support.HasValue || rate < 0)
                    Fail("control_points", "Initial beam angles are required and planned dose rates must be nonnegative when present.");
                if (cp.Contains(DicomTag.IsocenterPosition))
                {
                    var next = Numbers(cp, DicomTag.IsocenterPosition, "isocenter");
                    if (next.Length != 3) Fail("isocenter", "Isocenter coordinates must contain three finite patient-coordinate values.");
                    if (isocenter != null && next.Where((value, axis) => Math.Abs(value - isocenter[axis]) > 0.001).Any())
                        Fail("unsupported_orientation", "Moving-isocenter delivery is not supported for target-projection enrichment.");
                    isocenter = next;
                }
                if (isocenter == null) Fail("isocenter", "Initial patient-coordinate isocenter is required.");
                double weight = Number(cp, DicomTag.CumulativeMetersetWeight, "meterset_weights");
                if (weight < lastWeight || weight < 0 || weight > finalWeight ||
                    (index == 0 && weight != 0) || (index == cps.Count - 1 && Math.Abs(weight - finalWeight) > 1e-6))
                    Fail("meterset_weights", "Cumulative meterset weights must start at zero, be monotonic, and end at the declared final weight.");
                lastWeight = weight;
                if (cp.Contains(DicomTag.BeamLimitingDevicePositionSequence))
                {
                    var changed = new HashSet<string>();
                    foreach (var position in Sequence(cp, DicomTag.BeamLimitingDevicePositionSequence, "leaf_positions"))
                    {
                        string type = Text(position, DicomTag.RTBeamLimitingDeviceType);
                        if (!changed.Add(type) || !definitions.Any(device => device.Type == type))
                            Fail("leaf_positions", "A control point references a duplicate or undefined beam-limiting device.");
                        positions[type] = Numbers(position, DicomTag.LeafJawPositions, "leaf_positions");
                    }
                }
                var geometry = BuildGeometry(definitions, positions, fixedBounds);
                if (fixedBounds && index > 0 && !SameRectangle(geometry.FixedBoundingBox, result.ControlPoints[0].Aperture.FixedBoundingBox))
                    Fail("fixed_bounds", "Declared dual-layer X/Y aperture limits must remain fixed at every control point; they are not assumed to be physical jaws.");
                result.ControlPoints.Add(new ReviewControlPointSample
                {
                    Index = index, GantryAngleDegrees = gantry.Value, CollimatorAngleDegrees = collimator.Value,
                    PatientSupportAngleDegrees = support.Value, PlannedDoseRateMuPerMin = rate,
                    CumulativeMetersetWeight = weight / finalWeight, IsocenterMm = (double[])isocenter.Clone(), Aperture = geometry
                });
            }
            return result;
        }

        private static List<DeviceDefinition> ReadDevices(DicomDataset beam)
        {
            var result = new List<DeviceDefinition>();
            foreach (var item in Sequence(beam, DicomTag.BeamLimitingDeviceSequence, "leaf_boundaries"))
            {
                string type = Text(item, DicomTag.RTBeamLimitingDeviceType);
                if (result.Any(device => device.Type == type)) Fail("leaf_boundaries", "Beam-limiting-device definitions are duplicated.");
                string axis = type == "MLCX" || type == "MLCX1" || type == "MLCX2" ? "X" : type == "MLCY" ? "Y" : null;
                if (axis == null && type != "X" && type != "ASYMX" && type != "Y" && type != "ASYMY")
                    Fail("leaf_boundaries", "The beam-limiting device type is unsupported; no geometry is guessed.");
                int pairs = Integer(item, DicomTag.NumberOfLeafJawPairs, "leaf_boundaries");
                if (pairs < 1 || pairs > 2000 || (axis == null && pairs != 1))
                    Fail("leaf_boundaries", "The declared number of leaf or jaw pairs is invalid.");
                double[] boundaries = axis == null ? null : Numbers(item, DicomTag.LeafPositionBoundaries, "leaf_boundaries");
                if (axis != null && (boundaries.Length != pairs + 1 ||
                    boundaries.Skip(1).Where((value, index) => value <= boundaries[index]).Any()))
                    Fail("leaf_boundaries", "Each physical MLC layer requires its own exact, strictly increasing N+1 leaf boundaries.");
                result.Add(new DeviceDefinition { Type = type, Axis = axis, Pairs = pairs, Boundaries = boundaries });
            }
            int xJaws = result.Count(device => device.Type == "X" || device.Type == "ASYMX");
            int yJaws = result.Count(device => device.Type == "Y" || device.Type == "ASYMY");
            bool firstLayer = result.Any(device => device.Type == "MLCX1"), secondLayer = result.Any(device => device.Type == "MLCX2");
            if (firstLayer != secondLayer || firstLayer && result.Any(device => device.Type == "MLCX" || device.Type == "MLCY"))
                Fail("leaf_boundaries", "The dual-layer extension requires both MLCX1 and MLCX2, without an ambiguous additional MLC definition.");
            if (xJaws > 1 || yJaws > 1 || xJaws != yJaws)
                Fail("leaf_boundaries", "Partial or duplicated jaw-axis geometry is unsupported; missing jaws are never synthesized.");
            return result;
        }

        private static ApertureGeometry BuildGeometry(List<DeviceDefinition> definitions, Dictionary<string, double[]> positions, bool fixedBounds)
        {
            var aperture = new ApertureGeometry(); double[] x = null, y = null;
            foreach (var device in definitions)
            {
                double[] values;
                if (!positions.TryGetValue(device.Type, out values) || values.Length != 2 * device.Pairs ||
                    Enumerable.Range(0, device.Pairs).Any(index => values[index] > values[index + device.Pairs]))
                    Fail("leaf_positions", "Every physical device requires complete ordered bank positions, present initially or inherited explicitly.");
                if (device.Axis == null)
                {
                    if (device.Type == "X" || device.Type == "ASYMX") x = values;
                    else y = values;
                }
                else aperture.Layers.Add(new ApertureLayer
                {
                    Label = device.Type, LeafTravelAxis = device.Axis, LeafBoundariesMm = (double[])device.Boundaries.Clone(),
                    Bank1PositionsMm = values.Take(device.Pairs).ToArray(), Bank2PositionsMm = values.Skip(device.Pairs).ToArray()
                });
            }
            if (x != null && y != null)
            {
                var limits = new ApertureRectangle(x[0], y[0], x[1], y[1]);
                if (fixedBounds) aperture.FixedBoundingBox = limits;
                else aperture.Jaws = limits;
            }
            return aperture;
        }

        private static void ValidateOrientation(DicomDataset cp, DicomDataset first, int index)
        {
            foreach (DicomTag tag in new[] { DicomTag.TableTopPitchAngle, DicomTag.TableTopRollAngle,
                DicomTag.TableTopEccentricAngle, DicomTag.GantryPitchAngle })
                if (cp.Contains(tag) && Math.Abs(Number(cp, tag, "unsupported_orientation")) > 0.0001)
                    Fail("unsupported_orientation", "Pitch, roll, eccentric, or gantry-pitch rotations are not supported for target-projection enrichment.");
            foreach (DicomTag tag in new[] { DicomTag.TableTopVerticalPosition, DicomTag.TableTopLongitudinalPosition, DicomTag.TableTopLateralPosition })
                if (index > 0 && cp.Contains(tag) && (!first.Contains(tag) ||
                    Math.Abs(Number(cp, tag, "unsupported_orientation") - Number(first, tag, "unsupported_orientation")) > 0.001))
                    Fail("unsupported_orientation", "Changing tabletop translations are not supported for target-projection enrichment.");
        }

        private static void ValidateBeamMatch(ReviewBeamAnalysis active, ReviewBeamAnalysis source)
        {
            if (!active.MetersetMu.HasValue || !Finite(active.MetersetMu.Value) ||
                Math.Abs(active.MetersetMu.Value - source.MetersetMu.Value) > Math.Max(0.01, source.MetersetMu.Value * 1e-6) ||
                active.ControlPoints == null || active.ControlPoints.Count != source.ControlPoints.Count)
                Fail("beam_match", "Treatment-beam MU or control point count differs from the active plan.");
            double last = active.ControlPoints.Last().CumulativeMetersetWeight;
            if (!Finite(last) || last <= 0) Fail("control_point_match", "Active-plan meterset weights are unavailable.");
            for (int index = 0; index < source.ControlPoints.Count; index++)
            {
                var a = active.ControlPoints[index]; var b = source.ControlPoints[index];
                if (a.Index != b.Index || !AnglesMatch(a.GantryAngleDegrees, b.GantryAngleDegrees) ||
                    !AnglesMatch(a.CollimatorAngleDegrees, b.CollimatorAngleDegrees) || !AnglesMatch(a.PatientSupportAngleDegrees, b.PatientSupportAngleDegrees) ||
                    !Finite(a.CumulativeMetersetWeight) || Math.Abs(a.CumulativeMetersetWeight / last - b.CumulativeMetersetWeight) > 1e-6 ||
                    a.IsocenterMm == null || a.IsocenterMm.Length != 3 ||
                    a.IsocenterMm.Where((value, axis) => !Finite(value) || Math.Abs(value - b.IsocenterMm[axis]) > 0.1).Any())
                    Fail("control_point_match", "Control point index, angles, normalized MU weight or patient-coordinate isocenter differs from the active plan.");
            }
        }

        private static bool AnglesMatch(double a, double b)
        {
            if (!Finite(a) || !Finite(b)) return false;
            double difference = Math.Abs((a - b) % 360); return Math.Min(difference, 360 - difference) <= 0.01;
        }
        private static bool SameRectangle(ApertureRectangle a, ApertureRectangle b)
        { return a != null && b != null && Math.Abs(a.X1-b.X1) <= 0.001 && Math.Abs(a.X2-b.X2) <= 0.001 && Math.Abs(a.Y1-b.Y1) <= 0.001 && Math.Abs(a.Y2-b.Y2) <= 0.001; }
        private static void SetProvenance(ReviewPlanAnalysis analysis)
        {
            analysis.GeometryProvenance = "DICOM RTPLAN: exact per-device leaf boundaries and CP bank positions, with CP inheritance; isocenter-plane mm. Dual MLCX1/MLCX2 X/Y tags are checked fixed clipping limits, not physical jaws. Planned geometry, not delivery reconstruction.";
            analysis.DoseRateProvenance = "DICOM DoseRateSet (300A,0115), inherited until changed, for the segment beginning at each CP. Missing values remain unavailable; not measured dose rate.";
        }
        private static void EnsureInitialized()
        {
            lock (InitializationLock)
            {
                if (_initialized) return;
                new DicomSetupBuilder().RegisterServices(services =>
                    services.AddFellowOakDicom().AddLogging(logging => logging.ClearProviders())).Build();
                _initialized = true;
            }
        }
        private static IList<DicomDataset> Sequence(DicomDataset dataset, DicomTag tag, string code)
        {
            DicomSequence sequence;
            if (!dataset.TryGetSequence(tag, out sequence) || sequence.Items.Count == 0 || sequence.Items.Count > 20000)
                Fail(code, "A required RTPLAN sequence is absent, empty, or exceeds the supported import size.");
            return sequence.Items;
        }
        private static string Text(DicomDataset dataset, DicomTag tag) { return dataset.GetSingleValueOrDefault<string>(tag, string.Empty).Trim(); }
        private static string RequiredUid(DicomDataset dataset, DicomTag tag)
        {
            string value = Text(dataset, tag);
            if (value.Length == 0 || value.Length > 64 || value.Split('.').Any(part => part.Length == 0 || part.Any(character => !char.IsDigit(character))))
                Fail("plan_identity", "Required DICOM identity fields are missing or malformed.");
            return value;
        }
        private static int Integer(DicomDataset dataset, DicomTag tag, string code)
        {
            int value; if (!dataset.TryGetSingleValue(tag, out value)) Fail(code, "A required RTPLAN integer attribute is missing or invalid."); return value;
        }
        private static double Number(DicomDataset dataset, DicomTag tag, string code)
        {
            double value; if (!dataset.TryGetSingleValue(tag, out value) || !Finite(value)) Fail(code, "A required RTPLAN numeric attribute is missing or invalid."); return value;
        }
        private static double[] Numbers(DicomDataset dataset, DicomTag tag, string code)
        {
            double[] values; if (!dataset.TryGetValues(tag, out values) || values.Length == 0 || !values.All(Finite))
                Fail(code, "A required RTPLAN numeric array is absent, empty, or non-finite."); return values;
        }
        private static double? InheritNumber(DicomDataset cp, DicomTag tag, double? previous)
        { return cp.Contains(tag) ? Number(cp, tag, "control_points") : previous; }
        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static void Fail(string code, string safeMessage) { throw new RtPlanImportException(code, safeMessage); }
        private sealed class DeviceDefinition
        { public string Type; public string Axis; public int Pairs; public double[] Boundaries; }
    }
}
