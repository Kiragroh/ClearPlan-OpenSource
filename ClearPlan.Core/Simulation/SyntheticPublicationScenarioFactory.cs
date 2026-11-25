using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using ClearPlan.Core.Fields;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Simulation
{
    public static class SyntheticPublicationScenarioFactory
    {
        private static readonly string[] PublicationIds = { "publication-single-layer", "publication-dual-layer" };
        public static IList<string> ScenarioIds { get { return Array.AsReadOnly(PublicationIds); } }

        public static ReviewSnapshot Create(string scenarioId)
        {
            if (string.Equals(scenarioId, PublicationIds[0], StringComparison.Ordinal)) return Create(false);
            if (string.Equals(scenarioId, PublicationIds[1], StringComparison.Ordinal)) return Create(true);
            throw new ArgumentException("Unknown publication scenario.", "scenarioId");
        }

        private const double PrescriptionGy = 60;
        private const double VoxelSizeMm = 3;
        private const double DoseBinGy = 0.1;
        private const int DoseBinCount = 651;
        private const string Watermark = "SYNTHETIC DEMONSTRATION — NOT FOR CLINICAL USE";
        private const string Provenance = "Deterministic mathematical ellipsoid phantom; no patient source data. " +
            "CT planes, structure contours, dose contours, DVHs and target-quality indices use the same analytical 3D definitions. " +
            "Volume sampling: 3 mm isotropic voxel centers; cumulative DVH bins: 0.1 Gy. " +
            "Analytical phantom dose is not dose calculated from apertures; this is a software demonstration, not physical dosimetric validation.";

        public static ReviewSnapshot Create(bool dualLayer)
        {
            var structures = Structures();
            string id = PublicationIds[dualLayer ? 1 : 0];
            var snapshot = new ReviewSnapshot
            {
                SchemaVersion = ReviewSnapshot.CurrentSchemaVersion,
                ScenarioId = id, ScenarioTitle = dualLayer ? "Synthetic jawless dual-layer review" : "Synthetic single-layer review",
                ScenarioDescription = "Coherent analytical phantom with illustrative aperture geometry and estimated dose rate.",
                Seed = dualLayer ? 2202 : 2201, Synthetic = true,
                GeneratedUtc = new DateTimeOffset(2026, 9, 11, 0, 0, 0, TimeSpan.Zero),
                PatientDisplayLabel = "Synthetic phantom — no patient data",
                PlanDisplayLabel = dualLayer ? "Synthetic dual-layer VMAT" : "Synthetic single-layer VMAT",
                ActivePlanKey = id, ProvenanceText = Provenance,
                Report = new ReviewReportMetadata
                {
                    Title = "ClearPlan synthetic review", Subtitle = "Analytical phantom · 30 × 2 Gy",
                    ModeLabel = Watermark, Watermark = Watermark, OutputFileLabel = id + ".pdf",
                    Notes = new List<string> { Provenance, "Demonstration goals are illustrative configuration examples, not clinical recommendations." }
                }
            };
            snapshot.Plans.Add(new ReviewPlanRow
            {
                PlanKey = id, DisplayLabel = snapshot.PlanDisplayLabel, CreatedUtc = snapshot.GeneratedUtc,
                DosePerFractionGy = 2, TotalDoseGy = PrescriptionGy, FractionCount = 30,
                TargetDisplayLabel = "PTV_60", Status = ReviewStatusCodes.Info
            });
            snapshot.Sources.Add(new ReviewSourceStatus
            {
                StableId = "publication-phantom", SourceCode = "analytical-phantom-v1", SourceType = "embedded",
                Status = ReviewStatusCodes.Available, PathDisplayLabel = "Public synthetic phantom v1", Message = Provenance
            });
            snapshot.Sources.Add(new ReviewSourceStatus
            {
                StableId = "publication-goals", SourceCode = "illustrative-goals-v1", SourceType = "embedded",
                Status = ReviewStatusCodes.Available, PathDisplayLabel = "Illustrative defaults v1",
                Message = "Synthetic example constraints; no institution-specific clinical database."
            });
            snapshot.DvhSeries = SampleDvh(structures);
            AddGoals(snapshot);
            snapshot.PlanImages = CreatePlanes(snapshot.ActivePlanKey, structures);
            snapshot.PlanAnalysis = CreatePlanAnalysis(dualLayer, snapshot, structures);
            AddFieldRows(snapshot);
            snapshot.PlanCheckRows.Add(new ReviewCheckRow
            {
                CheckCode = "synthetic-fractionation", Category = "Default", Message = "Synthetic prescription: 30 fractions × 2 Gy = 60 Gy.",
                ObservedValue = "30 × 2 Gy", ExpectedValue = "60 Gy", Unit = ReviewUnitCodes.Text,
                Status = ReviewStatusCodes.Pass, Severity = ReviewSeverityCodes.None
            });
            snapshot.PlanCheckRows.Add(new ReviewCheckRow
            {
                CheckCode = "synthetic-approval", Category = "Default", Message = "Injected synthetic finding: treatment approval is absent.",
                ObservedValue = "Unapproved", ExpectedValue = "Approved", Unit = ReviewUnitCodes.Text,
                Status = ReviewStatusCodes.Fail, Severity = ReviewSeverityCodes.Error
            });
            return snapshot;
        }

        // Positions/radii are millimeters in the declared synthetic HFS coordinate system.
        private static Ellipsoid[] Structures()
        {
            return new[]
            {
                new Ellipsoid("PTV_60", "#D1495B", 0, 0, 0, 40, 30, 55, "PTV"),
                new Ellipsoid("CTV_60", "#AF54C2", 0, 0, 0, 30, 23, 43, "CTV"),
                new Ellipsoid("SpinalCord", "#137D8D", 0, 55, 0, 9, 9, 80, null),
                new Ellipsoid("Parotid_L", "#C78B1E", 43, -15, 10, 20, 22, 35, null),
                new Ellipsoid("Parotid_R", "#3678B8", -62, -15, 10, 20, 22, 35, null),
                new Ellipsoid("External", "#67768C", 0, 0, 0, 150, 105, 180, null)
            };
        }

        private static double DoseGy(double targetRadiusSquared)
        {
            double radius = Math.Sqrt(targetRadiusSquared);
            return radius <= 1 ? 63 - 2.4 * radius * radius : 60.6 * Math.Exp(-3 * (radius - 1));
        }

        private static double PhantomHu(double x, double y, double z, Ellipsoid[] structures)
        {
            var body = structures[5];
            double radiusSquared = body.RadiusSquared(x, y, z);
            if (radiusSquared > 1) return -1000;
            if (radiusSquared > 0.9) return -85;
            if (structures[2].Contains(x, y, z)) return 50;
            if (x * x / 225 + (y - 55) * (y - 55) / 225 + z * z / 8100 < 1) return 850;
            if (structures[0].Contains(x, y, z)) return 85;
            if (structures[3].Contains(x, y, z) || structures[4].Contains(x, y, z)) return 55;
            return 30 + 8 * Math.Cos(x / 25) * Math.Cos(z / 40);
        }

        private static List<ReviewDvhSeries> SampleDvh(Ellipsoid[] structures)
        {
            var counts = structures.Select(s => new int[DoseBinCount]).ToArray();
            for (int z = -180; z <= 180; z += 3)
            for (int y = -105; y <= 105; y += 3)
            for (int x = -150; x <= 150; x += 3)
            {
                if (!structures[5].Contains(x, y, z)) continue;
                int bin = (int)Math.Floor(DoseGy(structures[0].RadiusSquared(x, y, z)) / DoseBinGy);
                for (int s = 0; s < structures.Length; s++)
                    if (structures[s].Contains(x, y, z)) counts[s][bin]++;
            }
            var result = new List<ReviewDvhSeries>();
            for (int s = 0; s < structures.Length; s++)
            {
                var definition = structures[s];
                int total = counts[s].Sum(), cumulative = 0;
                var points = new ReviewDvhPoint[DoseBinCount];
                for (int bin = DoseBinCount - 1; bin >= 0; bin--)
                {
                    cumulative += counts[s][bin];
                    points[bin] = new ReviewDvhPoint { DoseGy = bin * DoseBinGy, VolumePercent = 100.0 * cumulative / total };
                }
                result.Add(new ReviewDvhSeries
                {
                    StableId = "dvh-" + definition.Id, StructureId = definition.Id, DisplayName = definition.Id,
                    Role = definition.TargetKind != null ? ReviewDvhRoleCodes.Target : s == 5 ? ReviewDvhRoleCodes.External : ReviewDvhRoleCodes.OrganAtRisk,
                    ColorHex = definition.Color, LineStyle = ReviewLineStyleCodes.Solid, Selected = s != 5,
                    VolumeCc = total * VoxelSizeMm * VoxelSizeMm * VoxelSizeMm / 1000,
                    Points = points.ToList(), TargetKind = definition.TargetKind,
                    RequiredForTargetReview = definition.TargetKind != null,
                    TargetSelectionReason = definition.TargetKind == "CTV" ? "Largest contained target (CTV); explicitly defined inside the synthetic PTV." :
                        definition.TargetKind == "PTV" ? "Synthetic prescription target." : null
                });
            }
            return result;
        }

        private static void AddGoals(ReviewSnapshot snapshot)
        {
            AddGoal(snapshot, "PTV_60", "D98", ">=", 57, 55, "D98 ≥ 95% of the 60 Gy prescription.");
            AddGoal(snapshot, "PTV_60", "D2", "<=", 64.2, 66, "D2 is the dose to 2% of the target volume.");
            AddGoal(snapshot, "CTV_60", "D98", ">=", 57, 55, "D98 ≥ 95% of the 60 Gy prescription.");
            AddGoal(snapshot, "SpinalCord", "D2", "<=", 20, 25, "Illustrative near-maximum cord constraint.");
            AddGoal(snapshot, "Parotid_L", "Dmean", "<=", 26, 30, "Illustrative mean-dose constraint.");
            AddGoal(snapshot, "Parotid_L", "V30", "<=", 35, 45, "Illustrative volume receiving at least 30 Gy.");
            AddGoal(snapshot, "Parotid_R", "Dmean", "<=", 26, 30, "Illustrative mean-dose constraint.");
            foreach (var group in snapshot.PqmRows.GroupBy(r => r.ResolvedStructureId))
                snapshot.StructureMappings.Add(new ReviewStructureMapping
                {
                    StableId = "mapping-" + group.Key, TemplateStructure = group.Key, SelectedStructureId = group.Key,
                    AvailableStructureIds = snapshot.DvhSeries.Select(s => s.StructureId).ToList(),
                    Status = ReviewStatusCodes.Pass, Message = "Exact synthetic structure name."
                });
        }

        private static void AddGoal(ReviewSnapshot snapshot, string structure, string objective, string comparator,
            double goal, double variation, string explanation)
        {
            var curve = snapshot.DvhSeries.Single(s => s.StructureId == structure);
            double achieved = objective == "Dmean" ? SyntheticDvhMetrics.MeanDoseGy(curve.Points) :
                objective == "V30" ? SyntheticDvhMetrics.RelativeVolumePercentAtDoseGy(curve.Points, 30) :
                SyntheticDvhMetrics.DoseGyAtRelativeVolumePercent(curve.Points, objective == "D98" ? 98 : 2);
            string status = SyntheticDvhMetrics.EvaluateThresholdStatus(achieved, comparator, goal, variation);
            snapshot.PqmRows.Add(new ReviewPqmRow
            {
                StableId = "pqm-" + structure + "-" + objective, TemplateCode = structure + "-" + objective,
                TemplateStructure = structure, ResolvedStructureId = structure, StructureOptions = new List<string> { structure },
                Objective = objective, Comparator = comparator, Goal = goal, Variation = variation, AchievedValue = achieved,
                Unit = objective == "V30" ? ReviewUnitCodes.Percent : ReviewUnitCodes.Gray, Status = status,
                Severity = status == ReviewStatusCodes.Pass ? ReviewSeverityCodes.None : status == ReviewStatusCodes.Variation ? ReviewSeverityCodes.Warning : ReviewSeverityCodes.Error,
                Explanation = explanation + " Calculated from the displayed sampled cumulative DVH; synthetic demonstration threshold.",
                SourceLabel = "Illustrative defaults v1", MappingDescription = "Exact synthetic structure name"
            });
        }

        private static ReviewPlanAnalysis CreatePlanAnalysis(bool dualLayer, ReviewSnapshot snapshot, Ellipsoid[] structures)
        {
            var plan = SyntheticPlanAnalysisFactory.Create(dualLayer, true);
            plan.GeometryProvenance = "Illustrative synthetic apertures and perspective projection of the same PTV_60 ellipsoid, not a commissioned machine model. " +
                "The analytical phantom dose is not dose calculated from apertures. " + PlanAnalysisCalculator.SamplingDefinition;
            plan.DoseRateProvenance = "Estimated synthetic plan trace; calculated from cumulative MU, gantry angles and explicit synthetic speed limits. No measured delivery.";
            var targetMesh = CreateTargetMesh(structures[0]);
            var volume = CreateCtVolume(structures);
            var profile = new DoseRateEstimationProfile
            {
                Id = dualLayer ? "SYNTHETIC_DUAL_LAYER_RATE" : "SYNTHETIC_SINGLE_LAYER_RATE",
                MachineIds = new List<string> { dualLayer ? "SYNTHETIC_DUAL_LAYER" : "SYNTHETIC_C_ARM" },
                MaxGantrySpeedDegreesPerSecond = dualLayer ? 12 : 6,
                Assumption = "Public synthetic example only. Configured gantry-speed bound; excludes acceleration, leaf motion and holds."
            };
            foreach (var beam in plan.Beams)
            {
                beam.Technique = "VMAT"; beam.GantryDirection = "Clockwise";
                double total = 0;
                for (int i = 1; i < beam.ControlPoints.Count; i++)
                    total += IntervalWeight(i, beam.BeamNumber.Value);
                double cumulative = 0;
                foreach (var cp in beam.ControlPoints)
                {
                    if (cp.Index > 0) cumulative += IntervalWeight(cp.Index, beam.BeamNumber.Value);
                    cp.CumulativeMetersetWeight = cumulative / total;
                    cp.PlannedDoseRateMuPerMin = null;
                    var frame = SyntheticDrrFactory.CreateHfsFrame(cp.GantryAngleDegrees, cp.CollimatorAngleDegrees, 0, 1000);
                    cp.TargetOutlines.Clear();
                    cp.TargetProjectionStrips = MeshTargetProjector.Project(targetMesh, frame, 1.25);
                    cp.TargetProjectionResolutionMm = 1.25;
                    cp.TargetProjectionProvenance = "Same synthetic PTV_60 ellipsoid; triangulated surface and perspective projection at SAD 1000 mm, raster 1.25 mm.";
                    cp.IsocenterMm = new[] { 0.0, 0.0, 0.0 };
                    if (cp.Index == 0)
                        cp.BevImage = ProjectDrr(volume, cp);
                }
                DoseRateEstimator.Apply(beam, profile);
            }
            PlanAnalysisCalculator.Calculate(plan);
            var target = snapshot.DvhSeries.Single(s => s.StructureId == "PTV_60");
            var body = snapshot.DvhSeries.Single(s => s.StructureId == "External");
            var quality = TargetQualityCalculator.Calculate(target.StructureId, PrescriptionGy, target.VolumeCc,
                VolumeAtDose(target, PrescriptionGy), VolumeAtDose(body, PrescriptionGy), VolumeAtDose(body, PrescriptionGy / 2),
                SyntheticDvhMetrics.DoseGyAtRelativeVolumePercent(target.Points, 2), SyntheticDvhMetrics.DoseGyAtRelativeVolumePercent(target.Points, 98));
            quality.BodyStructureId = body.StructureId;
            quality.Note = (string.IsNullOrWhiteSpace(quality.Note) ? "" : quality.Note + " ") +
                "Synthetic analytical phantom; EXTERNAL whole-plan scope.";
            plan.TargetQuality.Add(quality);
            plan.TargetQualityNote = TargetQualityCalculator.Definition + " " + Provenance;
            return plan;
        }

        /// <summary>Regenerates any explicitly synthetic publication CP from the same phantom used by Create.</summary>
        public static BeamEyeViewImage CreateDrr(ReviewBeamAnalysis beam, ReviewControlPointSample cp)
        {
            if (beam == null || cp == null || beam.ControlPoints == null || !beam.ControlPoints.Contains(cp) ||
                (beam.MachineId != "SYNTHETIC_C_ARM" && beam.MachineId != "SYNTHETIC_DUAL_LAYER"))
                throw new ArgumentException("A control point belonging to an explicitly synthetic publication beam is required.");
            return ProjectDrr(CreateCtVolume(Structures()), cp);
        }

        private static BeamEyeViewImage ProjectDrr(CtVolume volume, ReviewControlPointSample cp)
        {
            var frame = SyntheticDrrFactory.CreateHfsFrame(cp.GantryAngleDegrees, cp.CollimatorAngleDegrees, cp.PatientSupportAngleDegrees, 1000);
            var image = CtDrrProjector.Project(volume, frame, 192, 140, CancellationToken.None);
            image.Synthetic = true; image.ControlPointIndex = cp.Index;
            image.GantryAngleDegrees = cp.GantryAngleDegrees; image.CollimatorAngleDegrees = cp.CollimatorAngleDegrees;
            image.BldToDisplayRotationDegrees = cp.CollimatorAngleDegrees % 360;
            image.PatientSupportAngleDegrees = cp.PatientSupportAngleDegrees;
            image.ProjectionDescription = "Control-point projection of the same analytical phantom as the three CT planes; synthetic HU-equivalent values. " +
                "Declared synthetic HFS frame, couch 0, SAD 1000 mm. " + CtDrrProjector.MethodDescription;
            return image;
        }

        private static CtVolume CreateCtVolume(Ellipsoid[] structures)
        {
            var volume = new CtVolume
            {
                SizeX = 101, SizeY = 71, SizeZ = 121, SpacingX = VoxelSizeMm, SpacingY = VoxelSizeMm, SpacingZ = VoxelSizeMm,
                Origin = new BeamPoint3D(-150, -105, -180), XAxis = new BeamPoint3D(1, 0, 0),
                YAxis = new BeamPoint3D(0, 1, 0), ZAxis = new BeamPoint3D(0, 0, 1), HounsfieldUnits = new float[101 * 71 * 121]
            };
            for (int z = 0; z < volume.SizeZ; z++)
            for (int y = 0; y < volume.SizeY; y++)
            for (int x = 0; x < volume.SizeX; x++)
                volume.HounsfieldUnits[(z * volume.SizeY + y) * volume.SizeX + x] =
                    (float)PhantomHu(-150 + x * VoxelSizeMm, -105 + y * VoxelSizeMm, -180 + z * VoxelSizeMm, structures);
            return volume;
        }

        private static TargetMeshGeometry CreateTargetMesh(Ellipsoid target)
        {
            const int rings = 24, sectors = 48;
            var vertices = new List<BeamPoint3D>();
            var triangles = new List<int>();
            for (int ring = 0; ring <= rings; ring++)
            for (int sector = 0; sector < sectors; sector++)
            {
                double latitude = ring * Math.PI / rings, longitude = sector * 2 * Math.PI / sectors;
                vertices.Add(new BeamPoint3D(
                    target.Center[0] + target.Radii[0] * Math.Sin(latitude) * Math.Cos(longitude),
                    target.Center[1] + target.Radii[1] * Math.Sin(latitude) * Math.Sin(longitude),
                    target.Center[2] + target.Radii[2] * Math.Cos(latitude)));
                if (ring == rings) continue;
                int a = ring * sectors + sector, b = ring * sectors + (sector + 1) % sectors;
                triangles.AddRange(new[] { a, b, a + sectors, b, b + sectors, a + sectors });
            }
            return new TargetMeshGeometry { Vertices = vertices, TriangleIndices = triangles.ToArray() };
        }

        private static double IntervalWeight(int index, int beamNumber)
        { return 0.04 + Math.Pow(Math.Sin((index - 0.5) * Math.PI / 30 + (beamNumber - 1) * 0.65), 2); }

        private static double? VolumeAtDose(ReviewDvhSeries curve, double doseGy)
        { return curve.VolumeCc * SyntheticDvhMetrics.RelativeVolumePercentAtDoseGy(curve.Points, doseGy) / 100; }

        private static void AddFieldRows(ReviewSnapshot snapshot)
        {
            var suggestions = FieldNameSuggester.Suggest("SYN", snapshot.PlanAnalysis.Beams.Select((beam, index) => new BeamNamingInput
            {
                CurrentId = beam.BeamId, BeamNumber = beam.BeamNumber.Value, TreatmentOrderIndex = index,
                GantryStartAngle = beam.ControlPoints.First().GantryAngleDegrees,
                GantryStopAngle = beam.ControlPoints.Last().GantryAngleDegrees, GantryDirection = "CW", PatientSupportAngle = 0
            }).ToList());
            foreach (var suggestion in suggestions)
            {
                snapshot.PlanAnalysis.Beams[suggestion.DisplayOrder - 1].BeamId = suggestion.ExpectedId;
                snapshot.FieldRows.Add(new ReviewFieldRow
                {
                    StableId = "field-" + suggestion.DisplayOrder, BeamNumber = suggestion.BeamNumber,
                    TreatmentOrder = suggestion.DisplayOrder, CurrentId = suggestion.ExpectedId, ExpectedId = suggestion.ExpectedId,
                    CurrentName = suggestion.SuggestedName, SuggestedName = suggestion.SuggestedName,
                    IdStatus = ReviewStatusCodes.Pass, NameStatus = ReviewStatusCodes.Pass
                });
            }
        }

        private static List<ReviewPlanImage> CreatePlanes(string planKey, Ellipsoid[] structures)
        {
            const int size = 401;
            const double spacing = 1;
            double center = (size - 1) / 2.0;
            var result = new List<ReviewPlanImage>();
            foreach (string kind in new[] { "transversal", "coronal", "sagittal" })
            {
                var plane = new PlanImagePlane(kind, size, size, kind == "sagittal" ? 0 : -center,
                    kind == "coronal" ? 0 : -center, kind == "transversal" ? 0 : center, spacing, spacing);
                var image = new ReviewPlanImage
                {
                    PlanKey = planKey, Kind = kind, Title = kind == "transversal" ? "Transversal" : kind == "coronal" ? "Coronal (orthogonal)" : "Sagittal",
                    Caption = "SYNTHETIC analytical phantom at isocenter; W400 / L40 HU-equivalent. Same 3D analytical dose as the DVH, not an aperture dose calculation.",
                    Synthetic = true, SourceStatus = ReviewStatusCodes.Available, WidthPixels = size, HeightPixels = size,
                    GrayscalePixels = new byte[size * size], PixelSpacingXMillimeters = spacing, PixelSpacingYMillimeters = spacing,
                    IsocenterPixelX = center, IsocenterPixelY = center,
                    LeftOrientation = kind == "sagittal" ? "A" : "R", RightOrientation = kind == "sagittal" ? "P" : "L",
                    TopOrientation = kind == "transversal" ? "A" : "S", BottomOrientation = kind == "transversal" ? "P" : "I",
                    OverlaySummary = "Analytical ellipsoid intersections and isodoses; sampled from the same phantom dose function as all displayed DVHs."
                };
                image.DosePlane = new ReviewImageDosePlane { Columns = size, Rows = size, PrescriptionGy = PrescriptionGy,
                    SamplesGy = new double[size * size], Source = "Synthetic analytical phantom dose sampled on the displayed plane" };
                double minX = size, maxX = 0, minY = size, maxY = 0;
                for (int row = 0; row < size; row++)
                for (int column = 0; column < size; column++)
                {
                    var world = plane.ToWorld(column, row);
                    int index = row * size + column;
                    image.GrayscalePixels[index] = OrthogonalImageGeometry.WindowHu(PhantomHu(world[0], world[1], world[2], structures), 40, 400);
                    double dose = DoseGy(structures[0].RadiusSquared(world[0], world[1], world[2]));
                    image.DosePlane.SamplesGy[index] = dose;
                    if (dose >= PrescriptionGy * 0.02)
                    { minX = Math.Min(minX, column); maxX = Math.Max(maxX, column); minY = Math.Min(minY, row); maxY = Math.Max(maxY, row); }
                }
                image.DoseFocusRegion = new ReviewImageDoseRegion
                {
                    PrescriptionPercent = 2, ThresholdGy = PrescriptionGy * 0.02,
                    MinPixelX = minX, MaxPixelX = maxX, MinPixelY = minY, MaxPixelY = maxY, CoverageLimited = false
                };
                double[] percentages = { 2, 20, 50, 80, 95, 100, 102 };
                string[] colors = { "#83A4B8", "#5795CB", "#27A3A1", "#71B650", "#E1BF30", "#ED8544", "#D75055" };
                for (int i = 0; i < percentages.Length; i++)
                {
                    double dose = percentages[i] * PrescriptionGy / 100;
                    image.Overlays.Add(new ReviewImageOverlay
                    {
                        Kind = "isodose", Label = percentages[i].ToString(CultureInfo.InvariantCulture) + "% Rx / " + dose.ToString("0.##", CultureInfo.InvariantCulture) + " Gy",
                        ColorHex = colors[i], DoseGy = dose, Source = "Analytical phantom dose", SourceStatus = ReviewStatusCodes.Available,
                        // This phantom is analytically radial in PTV-normalized coordinates. Use the
                        // exact level-set ellipse, avoiding linear interpolation across its dose kink.
                        Paths = EllipsoidContours(plane, structures[0], dose >= 60.6 ?
                            Math.Sqrt((63 - dose) / 2.4) : 1 - Math.Log(dose / 60.6) / 3)
                    });
                }
                foreach (var structure in structures)
                {
                    var contours = EllipsoidContours(plane, structure, 1);
                    if (contours.Count == 0) continue;
                    image.Overlays.Add(new ReviewImageOverlay
                    {
                        Kind = "structure", Label = structure.Id, ColorHex = structure.Color, Source = "Analytical ellipsoid intersection",
                        SourceStatus = ReviewStatusCodes.Available, Paths = contours
                    });
                }
                result.Add(image);
            }
            return result;
        }

        private static List<ReviewImagePath> EllipsoidContours(PlanImagePlane plane, Ellipsoid structure, double radialScale)
        {
            double offset = structure.Center[plane.NormalAxis] / (structure.Radii[plane.NormalAxis] * radialScale);
            if (Math.Abs(offset) >= 1) return new List<ReviewImagePath>();
            double scale = radialScale * Math.Sqrt(1 - offset * offset);
            var contour = new double[129][];
            for (int i = 0; i < contour.Length; i++)
            {
                double angle = i * 2 * Math.PI / (contour.Length - 1);
                var point = (double[])structure.Center.Clone(); point[plane.NormalAxis] = 0;
                point[plane.HorizontalAxis] += structure.Radii[plane.HorizontalAxis] * scale * Math.Cos(angle);
                point[plane.VerticalAxis] += structure.Radii[plane.VerticalAxis] * scale * Math.Sin(angle);
                contour[i] = point;
            }
            return PlanImageOverlayGeometry.ProjectContours(plane, new[] { contour });
        }

        private sealed class Ellipsoid
        {
            public readonly string Id, Color, TargetKind;
            public readonly double[] Center, Radii;
            public Ellipsoid(string id, string color, double x, double y, double z, double rx, double ry, double rz, string targetKind)
            { Id = id; Color = color; Center = new[] { x, y, z }; Radii = new[] { rx, ry, rz }; TargetKind = targetKind; }
            public bool Contains(double x, double y, double z) { return RadiusSquared(x, y, z) <= 1; }
            public double RadiusSquared(double x, double y, double z)
            {
                x -= Center[0]; y -= Center[1]; z -= Center[2];
                return x * x / (Radii[0] * Radii[0]) + y * y / (Radii[1] * Radii[1]) + z * z / (Radii[2] * Radii[2]);
            }
        }
    }
}
