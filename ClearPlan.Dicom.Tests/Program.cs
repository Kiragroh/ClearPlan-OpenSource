using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Dicom;
using FellowOakDicom;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace ClearPlan.Dicom.Tests
{
    internal static class Program
    {
        private static string _directory;
        private static int Main(string[] args)
        {
            if (args.Length > 0) return Inspect(args);
            new DicomSetupBuilder().RegisterServices(services =>
                services.AddFellowOakDicom().AddLogging(logging => logging.ClearProviders())).Build();
            _directory = Path.Combine(Path.GetTempPath(), "ClearPlanDicomTests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_directory);
            var tests = new Dictionary<string, Action>
            {
                { "DualLayerInheritance", DualLayerInheritance }, { "RealJaws", RealJaws },
                { "MissingBoundaries", MissingBoundaries }, { "ReversedGeometry", ReversedGeometry },
                { "AmbiguousFractionGroups", AmbiguousFractionGroups }, { "UnsupportedSop", UnsupportedSop },
                { "PrivateVarianSop", PrivateVarianSop }, { "MalformedSafeError", MalformedSafeError },
                { "NonMonotonicWeights", NonMonotonicWeights }, { "MissingRate", MissingRate },
                { "MatchingAndTargetRetention", MatchingAndTargetRetention }, { "MismatchIsAtomic", MismatchIsAtomic },
                { "IdentifiersStayInMemory", IdentifiersStayInMemory }, { "UnsupportedPitch", UnsupportedPitch },
                { "DualLayerFixedBounds", DualLayerFixedBounds }, { "MovingVirtualBounds", MovingVirtualBounds },
                { "MissingInitialPositions", MissingInitialPositions }, { "MatchGeometryAndMu", MatchGeometryAndMu },
                { "InspectionOutputIsIdentifierFree", InspectionOutputIsIdentifierFree }, { "PartialDualLayer", PartialDualLayer }
            };
            int failures = 0;
            try
            {
                foreach (var test in tests)
                {
                    try { test.Value(); Console.WriteLine("PASS RTPLAN." + test.Key); }
                    catch (Exception error) { failures++; Console.WriteLine("FAIL RTPLAN." + test.Key + ": " + error.Message); }
                }
            }
            finally { Directory.Delete(_directory, true); }
            Console.WriteLine(tests.Count + " RTPLAN tests, " + failures + " failures");
            return failures == 0 ? 0 : 1;
        }

        private static void DualLayerInheritance()
        {
            var plan = Read(Fixture());
            var beam = plan.Analysis.Beams.Single();
            Equal(200, beam.MetersetMu.Value); Equal(2, beam.MlcLayerCount);
            Check(beam.HasJaws == false, "Jawless geometry must stay jawless.");
            var cp = beam.ControlPoints;
            Equal(0.5, cp[1].CumulativeMetersetWeight); Equal(100, cp[1].IncrementalMetersetMu.Value);
            Equal(600, cp[1].PlannedDoseRateMuPerMin.Value); Equal(300, cp[2].PlannedDoseRateMuPerMin.Value);
            Equal(10, cp[2].GantryAngleDegrees); Equal(15, cp[2].CollimatorAngleDegrees);
            Equal(-15, cp[0].Aperture.Layers[1].LeafBoundariesMm[0]);
            Equal(-4, cp[1].Aperture.Layers[0].Bank1PositionsMm[0]);
            Equal(-4, cp[2].Aperture.Layers[0].Bank1PositionsMm[0]);
            Check(!ReferenceEquals(cp[1].Aperture.Layers[0].Bank1PositionsMm, cp[2].Aperture.Layers[0].Bank1PositionsMm), "CP geometry must be detached.");
        }
        private static void RealJaws()
        {
            var root = Fixture(); var beam = Beam(root);
            beam.GetSequence(DicomTag.BeamLimitingDeviceSequence).Items.RemoveAt(1);
            beam.GetSequence(DicomTag.BeamLimitingDeviceSequence).Items[0].AddOrUpdate(DicomTag.RTBeamLimitingDeviceType, "MLCX");
            Cp(root, 0).GetSequence(DicomTag.BeamLimitingDevicePositionSequence).Items.RemoveAt(1);
            for (int index = 0; index < 2; index++) Cp(root, index).GetSequence(DicomTag.BeamLimitingDevicePositionSequence)
                .Items[0].AddOrUpdate(DicomTag.RTBeamLimitingDeviceType, "MLCX");
            beam.GetSequence(DicomTag.BeamLimitingDeviceSequence).Items.Add(Device("ASYMX", 1, null));
            beam.GetSequence(DicomTag.BeamLimitingDeviceSequence).Items.Add(Device("ASYMY", 1, null));
            var positions = Cp(root, 0).GetSequence(DicomTag.BeamLimitingDevicePositionSequence).Items;
            positions.Add(Position("ASYMX", -7, 7)); positions.Add(Position("ASYMY", -8, 8));
            var result = Read(root).Analysis.Beams.Single();
            Check(result.HasJaws == true, "Real jaws missing."); Equal(-7, result.ControlPoints[2].Aperture.Jaws.X1);
        }
        private static void DualLayerFixedBounds()
        {
            var parsed = Read(FixedBoundsFixture()); var beam = parsed.Analysis.Beams.Single();
            Check(beam.HasJaws == false, "Declared dual-layer limits must not be reported as physical jaws.");
            Check(beam.ControlPoints[0].Aperture.Jaws == null, "Virtual limits must be separate from jaws.");
            Equal(-3, beam.ControlPoints[2].Aperture.FixedBoundingBox.X1);
            Equal(0.96, beam.ControlPoints[0].ApertureAreaCm2.Value);
            var merged = RtPlanReader.MatchAndApply(Active(parsed), parsed, parsed.SopInstanceUid);
            Equal(-3, merged.Beams[0].ControlPoints[2].Aperture.FixedBoundingBox.X1);
            merged.Beams[0].ControlPoints[2].Aperture.FixedBoundingBox.X1 = -2;
            Equal(-3, parsed.Analysis.Beams[0].ControlPoints[2].Aperture.FixedBoundingBox.X1);
        }
        private static void MovingVirtualBounds()
        {
            var root = FixedBoundsFixture(); Cp(root, 1).GetSequence(DicomTag.BeamLimitingDevicePositionSequence)
                .Items.Add(Position("X", -2, 3));
            Reject("fixed_bounds", () => Read(root));
        }
        private static void MissingInitialPositions()
        {
            var root = Fixture(); Cp(root, 0).GetSequence(DicomTag.BeamLimitingDevicePositionSequence).Items.RemoveAt(1);
            Reject("leaf_positions", () => Read(root));
        }
        private static void PartialDualLayer()
        {
            var root = Fixture(); Beam(root).GetSequence(DicomTag.BeamLimitingDeviceSequence).Items.RemoveAt(1);
            Reject("leaf_boundaries", () => Read(root));
        }
        private static void MatchGeometryAndMu()
        {
            var parsed = Read(Fixture()); var active = Active(parsed); active.Beams[0].MetersetMu += 1;
            Reject("beam_match", () => RtPlanReader.MatchAndApply(active, parsed, parsed.SopInstanceUid));
            active = Active(parsed); active.Beams[0].ControlPoints[2].IsocenterMm[1] += 1;
            Reject("control_point_match", () => RtPlanReader.MatchAndApply(active, parsed, parsed.SopInstanceUid));
            active = Active(parsed); active.Beams[0].BeamNumber = 2;
            Reject("beam_match", () => RtPlanReader.MatchAndApply(active, parsed, parsed.SopInstanceUid));
        }
        private static void InspectionOutputIsIdentifierFree()
        {
            var parsed = Read(FixedBoundsFixture()); string summary = Describe(parsed);
            Check(!summary.Contains(parsed.SopInstanceUid) && !summary.Contains(parsed.FrameOfReferenceUid) &&
                !summary.Contains(parsed.StudyInstanceUid) && !summary.Contains(_directory), "Inspection must not disclose identifiers or paths.");
            Check(summary.Contains("Physical jaws: no; fixed limits: yes"), "Inspection must distinguish limits from physical jaws.");
        }
        private static int Inspect(string[] args)
        {
            int selectedGroup = 0;
            if (args.Length != 2 && args.Length != 4 || args[0] != "--inspect" ||
                args.Length == 4 && (args[2] != "--fraction-group" || !int.TryParse(args[3], out selectedGroup) || selectedGroup <= 0))
            {
                Console.WriteLine("Usage: ClearPlan.Dicom.Tests.exe --inspect <RTPLAN path> [--fraction-group <number>]");
                return 2;
            }
            try
            {
                // Read-only diagnostics: no fixtures, logs, identifiers, or output artifacts are written.
                Console.Write(Describe(RtPlanReader.Read(args[1], args.Length == 4 ? (int?)selectedGroup : null)));
                return 0;
            }
            catch (RtPlanImportException error) { Console.WriteLine("RTPLAN inspection rejected: " + error.Code); return 1; }
            catch (Exception) { Console.WriteLine("RTPLAN inspection failed: dependency_or_runtime_failure"); return 1; }
        }
        private static string Describe(ParsedRtPlan parsed)
        {
            var output = new StringBuilder();
            output.AppendLine("Read-only RTPLAN inspection accepted; this does not validate active-plan matching.");
            output.AppendLine("SOP class: " + (parsed.SopClassUid == RtPlanReader.StandardRtPlanSopClass ? "standard RT Plan Storage" : "documented Varian private RT Plan Storage"));
            output.AppendLine("Treatment beam count: " + parsed.Analysis.Beams.Count);
            int index = 0;
            foreach (var beam in parsed.Analysis.Beams)
            {
                output.AppendLine("Beam ordinal " + (++index) + ": control point count " + beam.ControlPoints.Count);
                output.AppendLine("Physical jaws: " + (beam.HasJaws == true ? "yes" : "no") + "; fixed limits: " +
                    (beam.ControlPoints[0].Aperture.FixedBoundingBox == null ? "no" : "yes"));
                foreach (var layer in beam.ControlPoints[0].Aperture.Layers)
                    output.AppendLine("Layer " + layer.Label + ": pair count " + layer.Bank1PositionsMm.Length + "; boundary count " + layer.LeafBoundariesMm.Length);
                output.AppendLine("Planned dose-rate availability: " + beam.ControlPoints.Count(cp => cp.PlannedDoseRateMuPerMin.HasValue) + "/" + beam.ControlPoints.Count + " control points");
                output.AppendLine("Aperture geometry availability: " + beam.ControlPoints.Count(cp => cp.ApertureAreaCm2.HasValue) + "/" + beam.ControlPoints.Count + " control points");
            }
            return output.ToString();
        }
        private static void MissingBoundaries()
        {
            var root = Fixture(); Beam(root).GetSequence(DicomTag.BeamLimitingDeviceSequence).Items[1].Remove(DicomTag.LeafPositionBoundaries);
            Reject("leaf_boundaries", () => Read(root));
        }
        private static void ReversedGeometry()
        {
            var root = Fixture(); Cp(root, 0).GetSequence(DicomTag.BeamLimitingDevicePositionSequence).Items[0]
                .AddOrUpdate(DicomTag.LeafJawPositions, 10d, -10d, -10d, 10d);
            Reject("leaf_positions", () => Read(root));
            root = Fixture(); Beam(root).GetSequence(DicomTag.BeamLimitingDeviceSequence).Items[0]
                .AddOrUpdate(DicomTag.LeafPositionBoundaries, 10d, 0d, -10d);
            Reject("leaf_boundaries", () => Read(root));
        }
        private static void AmbiguousFractionGroups()
        {
            var root = Fixture(); root.GetSequence(DicomTag.FractionGroupSequence).Items.Add(Group(2, 150));
            Reject("fraction_group", () => Read(root));
            Equal(150, Read(root, 2).Analysis.Beams.Single().MetersetMu.Value);
            Reject("fraction_group", () => Read(root, 8));
        }
        private static void UnsupportedSop()
        {
            var root = Fixture(); root.AddOrUpdate(DicomTag.SOPClassUID, DicomUID.CTImageStorage.UID);
            Reject("sop_class", () => Read(root));
        }
        private static void PrivateVarianSop()
        {
            var root = Fixture(); root.AddOrUpdate(DicomTag.SOPClassUID, "1.2.246.352.70.1.70");
            Equal(1, Read(root).Analysis.Beams.Count);
        }
        private static void MalformedSafeError()
        {
            string path = Path.Combine(_directory, "malformed.dcm"); File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
            var error = Reject("dicom_read", () => RtPlanReader.Read(path));
            Check(!error.ToString().Contains(_directory), "Errors must not disclose file paths.");
            Check(error.InnerException == null, "Raw parser exceptions must not escape.");
        }
        private static void NonMonotonicWeights()
        {
            var root = Fixture(); Cp(root, 1).AddOrUpdate(DicomTag.CumulativeMetersetWeight, 120d);
            Reject("meterset_weights", () => Read(root));
            root = Fixture(); Beam(root).AddOrUpdate(DicomTag.FinalCumulativeMetersetWeight, 99d);
            Reject("meterset_weights", () => Read(root));
        }
        private static void MissingRate()
        {
            var root = Fixture(); Cp(root, 0).Remove(DicomTag.DoseRateSet); Cp(root, 2).Remove(DicomTag.DoseRateSet);
            Check(Read(root).Analysis.Beams.Single().ControlPoints.All(cp => cp.PlannedDoseRateMuPerMin == null), "Missing rate must not be invented.");
        }
        private static void MatchingAndTargetRetention()
        {
            var parsed = Read(Fixture()); var active = Active(parsed);
            var merged = RtPlanReader.MatchAndApply(active, parsed, parsed.SopInstanceUid, parsed.FrameOfReferenceUid);
            Equal(2, merged.DosePerFractionGy.Value); Equal(1, merged.Beams[0].ControlPoints[0].TargetOutlines.Count);
            Equal(1, merged.Beams[0].ControlPoints[0].TargetProjectionStrips.Count);
            Check(merged.Beams[0].ControlPoints[0].Aperture != null, "Geometry missing after match.");
            Check(active.Beams[0].ControlPoints[0].Aperture == null, "Active analysis must not mutate.");
            merged.Beams[0].ControlPoints[0].TargetOutlines[0][0].X = 99;
            Equal(-8, active.Beams[0].ControlPoints[0].TargetOutlines[0][0].X);
        }
        private static void MismatchIsAtomic()
        {
            var parsed = Read(Fixture()); var active = Active(parsed);
            Reject("plan_identity", () => RtPlanReader.MatchAndApply(active, parsed, DicomUID.Generate().UID));
            Reject("frame_identity", () => RtPlanReader.MatchAndApply(active, parsed, parsed.SopInstanceUid, DicomUID.Generate().UID));
            active.Beams[0].ControlPoints[1].GantryAngleDegrees += 1;
            Reject("control_point_match", () => RtPlanReader.MatchAndApply(active, parsed, parsed.SopInstanceUid));
            Check(active.Beams[0].ControlPoints[0].Aperture == null, "A failed match must not partially enrich.");
        }
        private static void IdentifiersStayInMemory()
        {
            var parsed = Read(Fixture()); string json = JsonConvert.SerializeObject(parsed);
            Check(!json.Contains(parsed.SopInstanceUid) && !json.Contains(parsed.StudyInstanceUid) && !json.Contains(parsed.FrameOfReferenceUid), "DICOM identifiers must not serialize.");
        }
        private static void UnsupportedPitch()
        {
            var root = Fixture(); Cp(root, 1).AddOrUpdate(DicomTag.TableTopPitchAngle, 2f);
            Reject("unsupported_orientation", () => Read(root));
        }

        private static ReviewPlanAnalysis Active(ParsedRtPlan parsed)
        {
            var active = JsonConvert.DeserializeObject<ReviewPlanAnalysis>(JsonConvert.SerializeObject(parsed.Analysis));
            active.DosePerFractionGy = 2; active.TargetStructureId = "PTV";
            for (int i = 0; i < active.Beams[0].ControlPoints.Count; i++)
            {
                var cp = active.Beams[0].ControlPoints[i]; cp.IsocenterMm = new[] { 1d, 2d, 3d };
                cp.TargetOutlines.Add(new List<BeamPoint> { new BeamPoint(-8,-8), new BeamPoint(8,-8), new BeamPoint(8,8), new BeamPoint(-8,8) });
                cp.TargetProjectionStrips.Add(new ApertureRectangle(-8,-8,8,8));
            }
            return active;
        }
        private static ParsedRtPlan Read(DicomDataset dataset, int? group = null)
        {
            string path = Path.Combine(_directory, Guid.NewGuid().ToString("N") + ".dcm");
            new DicomFile(dataset).Save(path);
            return RtPlanReader.Read(path, group);
        }
        private static DicomDataset Fixture()
        {
            var root = new DicomDataset();
            root.Add(DicomTag.SOPClassUID, DicomUID.RTPlanStorage.UID); root.Add(DicomTag.SOPInstanceUID, DicomUID.Generate().UID);
            root.Add(DicomTag.StudyInstanceUID, DicomUID.Generate().UID); root.Add(DicomTag.FrameOfReferenceUID, DicomUID.Generate().UID);
            root.Add(DicomTag.Modality, "RTPLAN"); root.Add(DicomTag.RTPlanGeometry, "PATIENT");
            root.Add(new DicomSequence(DicomTag.FractionGroupSequence, Group(1, 200)));
            var beam = new DicomDataset(); beam.Add(DicomTag.BeamNumber, 1); beam.Add(DicomTag.BeamName, "Beam-1");
            beam.Add(DicomTag.TreatmentDeliveryType, "TREATMENT"); beam.Add(DicomTag.BeamType, "DYNAMIC");
            beam.Add(DicomTag.RadiationType, "PHOTON"); beam.Add(DicomTag.PrimaryDosimeterUnit, "MU");
            beam.Add(DicomTag.NumberOfControlPoints, 3); beam.Add(DicomTag.FinalCumulativeMetersetWeight, 100d);
            beam.Add(new DicomSequence(DicomTag.BeamLimitingDeviceSequence,
                Device("MLCX1", 2, new[] { -10d, 0d, 10d }), Device("MLCX2", 3, new[] { -15d, -5d, 5d, 15d })));
            var cp0 = new DicomDataset(); cp0.Add(DicomTag.ControlPointIndex, 0); cp0.Add(DicomTag.CumulativeMetersetWeight, 0d);
            cp0.Add(DicomTag.GantryAngle, 0d); cp0.Add(DicomTag.BeamLimitingDeviceAngle, 15d); cp0.Add(DicomTag.PatientSupportAngle, 0d);
            cp0.Add(DicomTag.DoseRateSet, 600d); cp0.Add(DicomTag.IsocenterPosition, 1d, 2d, 3d);
            cp0.Add(new DicomSequence(DicomTag.BeamLimitingDevicePositionSequence,
                Position("MLCX1", -10, -10, 10, 10), Position("MLCX2", -8, -8, -8, 8, 8, 8)));
            var cp1 = new DicomDataset(); cp1.Add(DicomTag.ControlPointIndex, 1); cp1.Add(DicomTag.CumulativeMetersetWeight, 50d);
            cp1.Add(DicomTag.GantryAngle, 10d); cp1.Add(new DicomSequence(DicomTag.BeamLimitingDevicePositionSequence, Position("MLCX1", -4,-4,4,4)));
            var cp2 = new DicomDataset(); cp2.Add(DicomTag.ControlPointIndex, 2); cp2.Add(DicomTag.CumulativeMetersetWeight, 100d); cp2.Add(DicomTag.DoseRateSet, 300d);
            beam.Add(new DicomSequence(DicomTag.ControlPointSequence, cp0, cp1, cp2)); root.Add(new DicomSequence(DicomTag.BeamSequence, beam));
            return root;
        }
        private static DicomDataset FixedBoundsFixture()
        {
            var root = Fixture(); root.AddOrUpdate(DicomTag.SOPClassUID, RtPlanReader.VarianPrivateRtPlanSopClass);
            Beam(root).GetSequence(DicomTag.BeamLimitingDeviceSequence).Items.Add(Device("X", 1, null));
            Beam(root).GetSequence(DicomTag.BeamLimitingDeviceSequence).Items.Add(Device("Y", 1, null));
            var positions = Cp(root, 0).GetSequence(DicomTag.BeamLimitingDevicePositionSequence).Items;
            positions.Add(Position("X", -3, 3)); positions.Add(Position("Y", -8, 8));
            return root;
        }
        private static DicomDataset Group(int number, double mu)
        {
            var reference = new DicomDataset(); reference.Add(DicomTag.ReferencedBeamNumber, 1); reference.Add(DicomTag.BeamMeterset, mu);
            var group = new DicomDataset(); group.Add(DicomTag.FractionGroupNumber, number); group.Add(DicomTag.NumberOfBeams, 1);
            group.Add(new DicomSequence(DicomTag.ReferencedBeamSequence, reference)); return group;
        }
        private static DicomDataset Device(string type, int pairs, double[] boundaries)
        {
            var device = new DicomDataset(); device.Add(DicomTag.RTBeamLimitingDeviceType, type); device.Add(DicomTag.NumberOfLeafJawPairs, pairs);
            if (boundaries != null) device.Add(DicomTag.LeafPositionBoundaries, boundaries); return device;
        }
        private static DicomDataset Position(string type, params double[] positions)
        {
            var result = new DicomDataset(); result.Add(DicomTag.RTBeamLimitingDeviceType, type); result.Add(DicomTag.LeafJawPositions, positions); return result;
        }
        private static DicomDataset Beam(DicomDataset root) { return root.GetSequence(DicomTag.BeamSequence).Items[0]; }
        private static DicomDataset Cp(DicomDataset root, int index) { return Beam(root).GetSequence(DicomTag.ControlPointSequence).Items[index]; }
        private static RtPlanImportException Reject(string code, Action action)
        {
            try { action(); } catch (RtPlanImportException error) { Check(error.Code == code, "Expected safe code " + code + ", got " + error.Code); return error; }
            throw new Exception("Expected RTPLAN rejection: " + code);
        }
        private static void Equal(double expected, double actual) { Check(Math.Abs(expected-actual) < 1e-8, "Numeric assertion failed."); }
        private static void Check(bool value, string message) { if (!value) throw new Exception(message); }
    }
}
