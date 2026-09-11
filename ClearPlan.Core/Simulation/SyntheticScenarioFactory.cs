using System;
using System.Collections.Generic;
using System.Linq;
using ClearPlan.Core.Fields;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Simulation
{
    public static class SyntheticScenarioFactory
    {
        private const string SyntheticWatermark =
            "SYNTHETIC DEMONSTRATION — NOT FOR CLINICAL USE";

        private static readonly string[] OrderedScenarioIds =
        {
            "baseline-pass",
            "target-underdose",
            "oar-overdose",
            "metadata-plancheck",
            "field-and-mapping",
            "optional-path-fallback",
            "mixed-review"
        };

        private static readonly string[] RequiredStructures =
        {
            "PTV_60",
            "SpinalCord",
            "Parotid_L",
            "Parotid_R",
            "External"
        };

        public static IList<string> ScenarioIds
        {
            get { return Array.AsReadOnly(OrderedScenarioIds); }
        }

        public static IList<ReviewSnapshot> CreateAll()
        {
            return OrderedScenarioIds.Select(Create).ToList();
        }

        public static SyntheticScenarioExpectation GetExpectation(
            string scenarioId)
        {
            int index = Array.IndexOf(OrderedScenarioIds, scenarioId);
            if (index < 0)
            {
                throw new ArgumentException(
                    "Unknown synthetic scenario ID.",
                    "scenarioId");
            }

            return new SyntheticScenarioExpectation
            {
                ScenarioId = scenarioId,
                Seed = 1101 + index,
                SourceCount =
                    UsesOptionalFallback(scenarioId) ? 3 : 2,
                PqmRowCount = 6,
                PlanCheckRowCount = 8,
                FieldRowCount = 5,
                StructureMappingCount = 4,
                DvhSeriesCount = 5,
                RequiredStructureIds =
                    new List<string>(RequiredStructures)
            };
        }

        public static ReviewSnapshot Create(string scenarioId)
        {
            SyntheticScenarioExpectation expectation =
                GetExpectation(scenarioId);
            ReviewSnapshot snapshot = CreateBaseSnapshot(
                scenarioId,
                expectation.Seed);

            snapshot.DvhSeries = CreateDvhSeries(scenarioId);
            snapshot.PqmRows = CreatePqmRows(snapshot);
            snapshot.PlanCheckRows = CreatePlanCheckRows(scenarioId);
            snapshot.FieldRows = CreateFieldRows(
                UsesFieldAndMappingFindings(scenarioId));
            snapshot.StructureMappings = CreateMappings(
                UsesFieldAndMappingFindings(scenarioId));
            AddPlanAnalysis(snapshot);

            if (UsesOptionalFallback(scenarioId))
            {
                snapshot.Sources.Add(new ReviewSourceStatus
                {
                    StableId = "source-optional-reference",
                    SourceCode = "optional-reference",
                    SourceType = "optional",
                    Status = ReviewStatusCodes.Fallback,
                    Optional = true,
                    UsedFallback = true,
                    PathDisplayLabel = "Optional reference source",
                    Message =
                        "Optional source unavailable; embedded synthetic defaults are active."
                });
            }

            return snapshot;
        }

        public static void AddPlanAnalysis(ReviewSnapshot snapshot)
        {
            if (snapshot == null || !snapshot.Synthetic)
                throw new ArgumentException("Only explicitly synthetic snapshots can receive illustrative plan geometry.");
            snapshot.PlanAnalysis = PlanAnalysis.SyntheticPlanAnalysisFactory.Create(
                snapshot.ScenarioId == "mixed-review", snapshot.ScenarioId != "baseline-pass");
            snapshot.PlanAnalysis.TargetStructureId = "PTV_60";
            var target = snapshot.DvhSeries.FirstOrDefault(s => s.StructureId == "PTV_60");
            var body = snapshot.DvhSeries.FirstOrDefault(s => s.Role == ReviewDvhRoleCodes.External);
            if (target != null)
            {
                var stats = ReviewDvhStatisticsCalculator.Calculate(target.Points);
                var quality = PlanAnalysis.TargetQualityCalculator.Calculate(target.StructureId, 60, target.VolumeCc,
                    SyntheticVolumeAtDose(target, 60), SyntheticVolumeAtDose(body, 60), SyntheticVolumeAtDose(body, 30), stats.D2Gy, stats.D98Gy);
                quality.BodyStructureId = body == null ? null : body.StructureId;
                snapshot.PlanAnalysis.TargetQuality.Add(quality);
            }
            snapshot.PlanAnalysis.TargetQualityNote = PlanAnalysis.TargetQualityCalculator.Definition +
                " Synthetic: calculated from the displayed analytical target and External curves, not a physical dose calculation.";
            snapshot.PlanAnalysis.GeometryProvenance += " The projected target is an illustrative ellipse; the analytical DVH is not a dose calculation from this aperture geometry.";
        }

        private static double? SyntheticVolumeAtDose(ReviewDvhSeries series, double doseGy)
        {
            if (series == null || !series.VolumeCc.HasValue) return null;
            for (int i = 1; i < series.Points.Count; i++)
            {
                var a = series.Points[i - 1]; var b = series.Points[i];
                if (doseGy >= a.DoseGy && doseGy <= b.DoseGy && b.DoseGy > a.DoseGy)
                    return series.VolumeCc * (a.VolumePercent + (b.VolumePercent - a.VolumePercent) *
                        (doseGy - a.DoseGy) / (b.DoseGy - a.DoseGy)) / 100;
            }
            return null;
        }

        private static ReviewSnapshot CreateBaseSnapshot(
            string scenarioId,
            int seed)
        {
            string title;
            string description;
            GetScenarioText(scenarioId, out title, out description);
            DateTimeOffset generatedUtc =
                new DateTimeOffset(
                    2026,
                    7,
                    30,
                    8,
                    0,
                    0,
                    TimeSpan.Zero)
                .AddMinutes(seed - 1101);

            var snapshot = new ReviewSnapshot
            {
                SchemaVersion = ReviewSnapshot.CurrentSchemaVersion,
                ScenarioId = scenarioId,
                ScenarioTitle = title,
                ScenarioDescription = description,
                Seed = seed,
                Synthetic = true,
                GeneratedUtc = generatedUtc,
                PatientDisplayLabel = "Synthetic demonstration",
                PlanDisplayLabel = "Synthetic review plan",
                ProvenanceText =
                    "Analytically generated deterministic synthetic data; no clinical source data or identifiers were used.",
                ActivePlanKey = "synthetic-plan",
                Report = new ReviewReportMetadata
                {
                    Title = "ClearPlan synthetic review",
                    Subtitle = title,
                    ModeLabel = SyntheticWatermark,
                    Watermark = SyntheticWatermark,
                    OutputFileLabel =
                        "clearplan-" + scenarioId + "-report.pdf",
                    Notes = new List<string>
                    {
                        "Dose is expressed in Gy.",
                        "Cumulative volume is expressed as relative volume in percent."
                    }
                }
            };

            snapshot.Sources.Add(new ReviewSourceStatus
            {
                StableId = "source-scenario",
                SourceCode = "synthetic-scenario",
                SourceType = "embedded",
                Status = ReviewStatusCodes.Available,
                Optional = false,
                UsedFallback = false,
                PathDisplayLabel = "Embedded synthetic scenario",
                Message = "Deterministic scenario data loaded."
            });
            snapshot.Sources.Add(new ReviewSourceStatus
            {
                StableId = "source-constraints",
                SourceCode = "synthetic-constraints",
                SourceType = "embedded",
                Status = ReviewStatusCodes.Available,
                Optional = false,
                UsedFallback = false,
                PathDisplayLabel = "Embedded synthetic constraints",
                Message = "Synthetic review constraints loaded."
            });
            snapshot.Plans.Add(new ReviewPlanRow
            {
                PlanKey = "synthetic-plan",
                DisplayLabel = "Synthetic review plan",
                CreatedUtc = generatedUtc.AddDays(-1),
                DosePerFractionGy = 2.0,
                TotalDoseGy = 60.0,
                FractionCount = 30,
                TargetDisplayLabel = "PTV_60",
                Status = ReviewStatusCodes.Pass
            });
            return snapshot;
        }

        private static void GetScenarioText(
            string scenarioId,
            out string title,
            out string description)
        {
            switch (scenarioId)
            {
                case "baseline-pass":
                    title = "Baseline pass";
                    description =
                        "All synthetic PQM, PlanCheck, field, mapping, and DVH examples conform.";
                    return;
                case "target-underdose":
                    title = "Target underdose";
                    description =
                        "A shifted synthetic target DVH produces one variation and one failure.";
                    return;
                case "oar-overdose":
                    title = "Organ-at-risk overdose";
                    description =
                        "Shifted synthetic organ-at-risk DVHs produce one variation and one failure.";
                    return;
                case "metadata-plancheck":
                    title = "Metadata and technical checks";
                    description =
                        "Synthetic metadata and technical values demonstrate visible PlanCheck findings.";
                    return;
                case "field-and-mapping":
                    title = "Field naming and structure mapping";
                    description =
                        "Correct synthetic field IDs are shown with incorrect names and an ambiguous structure alias.";
                    return;
                case "optional-path-fallback":
                    title = "Optional source fallback";
                    description =
                        "A non-fatal optional source status uses embedded synthetic defaults.";
                    return;
                case "mixed-review":
                    title = "Integrated mixed review";
                    description =
                        "A publication-ready synthetic case combines target, organ-at-risk, metadata, field-name, and structure-mapping findings.";
                    return;
                default:
                    throw new ArgumentException(
                        "Unknown synthetic scenario ID.",
                        "scenarioId");
            }
        }

        private static List<ReviewPqmRow> CreatePqmRows(
            ReviewSnapshot snapshot)
        {
            var rows = new List<ReviewPqmRow>
            {
                Pqm(
                    "pqm-target-d95",
                    "target-d95",
                    "Target",
                    "PTV_60",
                    "D95",
                    ">=",
                    57.0,
                    55.0,
                    ReviewUnitCodes.Gray),
                Pqm(
                    "pqm-target-v95",
                    "target-v95",
                    "Target",
                    "PTV_60",
                    "V95",
                    ">=",
                    95.0,
                    92.0,
                    ReviewUnitCodes.Percent),
                Pqm(
                    "pqm-target-d2",
                    "target-d2",
                    "Target",
                    "PTV_60",
                    "D2",
                    "<=",
                    64.2,
                    66.0,
                    ReviewUnitCodes.Gray),
                Pqm(
                    "pqm-cord-dmax",
                    "cord-dmax",
                    "Spinal cord",
                    "SpinalCord",
                    "D0.1%",
                    "<=",
                    45.0,
                    48.0,
                    ReviewUnitCodes.Gray),
                Pqm(
                    "pqm-parotid-dmean",
                    "parotid-dmean",
                    "Parotid left",
                    "Parotid_L",
                    "Dmean",
                    "<=",
                    26.0,
                    36.0,
                    ReviewUnitCodes.Gray),
                Pqm(
                    "pqm-parotid-v30",
                    "parotid-v30",
                    "Parotid left",
                    "Parotid_L",
                    "V30",
                    "<=",
                    50.0,
                    60.0,
                    ReviewUnitCodes.Percent)
            };

            PopulatePqmValuesFromDvh(snapshot, rows);
            return rows;
        }

        private static ReviewPqmRow Pqm(
            string stableId,
            string templateCode,
            string templateStructure,
            string resolvedStructureId,
            string objective,
            string comparator,
            double goal,
            double variation,
            string unit)
        {
            return new ReviewPqmRow
            {
                StableId = stableId,
                TemplateCode = templateCode,
                TemplateStructure = templateStructure,
                ResolvedStructureId = resolvedStructureId,
                StructureOptions = new List<string>
                {
                    resolvedStructureId
                },
                Objective = objective,
                Comparator = comparator,
                Goal = goal,
                Variation = variation,
                Unit = unit,
                Status = ReviewStatusCodes.NotEvaluated,
                Severity = ReviewSeverityCodes.Info,
                Explanation = "Synthetic DVH metric has not been evaluated."
            };
        }

        private static void PopulatePqmValuesFromDvh(
            ReviewSnapshot snapshot,
            IEnumerable<ReviewPqmRow> rows)
        {
            ReviewPlanRow activePlan = snapshot.Plans.Single(plan =>
                plan.PlanKey == snapshot.ActivePlanKey);
            if (!activePlan.TotalDoseGy.HasValue)
            {
                throw new InvalidOperationException(
                    "Synthetic V95 evaluation requires total dose in Gy.");
            }

            foreach (ReviewPqmRow row in rows)
            {
                ReviewDvhSeries series = snapshot.DvhSeries.Single(item =>
                    item.StructureId == row.ResolvedStructureId);
                double rawValue;
                string metricExplanation;
                switch (row.Objective)
                {
                    case "D95":
                        rawValue =
                            SyntheticDvhMetrics.DoseGyAtRelativeVolumePercent(
                                series.Points,
                                95.0);
                        metricExplanation =
                            "D95 is obtained by inverse linear interpolation of the sampled relative DVH.";
                        break;
                    case "V95":
                        rawValue =
                            SyntheticDvhMetrics.RelativeVolumePercentAtDoseGy(
                                series.Points,
                                activePlan.TotalDoseGy.Value * 0.95);
                        metricExplanation =
                            "V95 is evaluated at 95 percent of the synthetic prescription dose.";
                        break;
                    case "D2":
                        rawValue =
                            SyntheticDvhMetrics.DoseGyAtRelativeVolumePercent(
                                series.Points,
                                2.0);
                        metricExplanation =
                            "D2 is obtained by inverse linear interpolation of the sampled relative DVH.";
                        break;
                    case "D0.1%":
                        rawValue =
                            SyntheticDvhMetrics.DoseGyAtRelativeVolumePercent(
                                series.Points,
                                0.1);
                        metricExplanation =
                            "D0.1 percent is a near-maximum dose surrogate obtained from the sampled relative DVH.";
                        break;
                    case "Dmean":
                        rawValue =
                            SyntheticDvhMetrics.MeanDoseGy(series.Points);
                        metricExplanation =
                            "Dmean is the trapezoidal integral of the sampled cumulative relative DVH.";
                        break;
                    case "V30":
                        rawValue =
                            SyntheticDvhMetrics.RelativeVolumePercentAtDoseGy(
                                series.Points,
                                30.0);
                        metricExplanation =
                            "V30 is obtained by linear interpolation at 30 Gy.";
                        break;
                    default:
                        throw new InvalidOperationException(
                            "Unsupported synthetic PQM objective: " +
                            row.Objective);
                }

                double achieved = Math.Round(
                    rawValue,
                    1,
                    MidpointRounding.AwayFromZero);
                if (!row.Goal.HasValue || !row.Variation.HasValue)
                {
                    throw new InvalidOperationException(
                        "Synthetic PQM thresholds must be explicit.");
                }

                row.AchievedValue = achieved;
                row.Status = SyntheticDvhMetrics.EvaluateThresholdStatus(
                    achieved,
                    row.Comparator,
                    row.Goal.Value,
                    row.Variation.Value);
                row.Severity = row.Status == ReviewStatusCodes.Pass
                    ? ReviewSeverityCodes.None
                    : row.Status == ReviewStatusCodes.Variation
                        ? ReviewSeverityCodes.Warning
                        : ReviewSeverityCodes.Error;
                row.Explanation = metricExplanation + " Result: " + row.Status + ".";
            }
        }

        private static List<ReviewCheckRow> CreatePlanCheckRows(
            string scenarioId)
        {
            var rows = new List<ReviewCheckRow>
            {
                Check(
                    "appointment-present",
                    "Metadata",
                    "Present",
                    "Present",
                    ReviewUnitCodes.Boolean,
                    "Synthetic appointment metadata is present."),
                Check(
                    "ct-age-days",
                    "Metadata",
                    "5",
                    "<= 14",
                    ReviewUnitCodes.Count,
                    "Synthetic image age is within the configured interval."),
                Check(
                    "user-origin-set",
                    "Geometry",
                    "Set",
                    "Set",
                    ReviewUnitCodes.Boolean,
                    "Synthetic user origin is configured."),
                Check(
                    "hu-table",
                    "Calculation",
                    "Synthetic default table",
                    "Synthetic default table",
                    ReviewUnitCodes.Text,
                    "Synthetic calibration table conforms."),
                Check(
                    "treatment-approval",
                    "Approval",
                    "Approved",
                    "Approved",
                    ReviewUnitCodes.Text,
                    "Synthetic treatment approval is present."),
                Check(
                    "calculation-grid",
                    "Calculation",
                    "2.5 mm",
                    "<= 3.0 mm",
                    ReviewUnitCodes.Text,
                    "Synthetic calculation grid conforms."),
                Check(
                    "treatment-unit",
                    "Technique",
                    "Synthetic treatment unit",
                    "Synthetic treatment unit",
                    ReviewUnitCodes.Text,
                    "Synthetic treatment unit conforms."),
                Check(
                    "fractionation",
                    "Prescription",
                    "30 x 2 Gy",
                    "30 x 2 Gy",
                    ReviewUnitCodes.Text,
                    "Synthetic fractionation conforms.")
            };

            if (scenarioId == "metadata-plancheck" ||
                scenarioId == "mixed-review")
            {
                string findingPrefix = scenarioId == "mixed-review"
                    ? "Injected synthetic finding: "
                    : string.Empty;
                SetCheckFinding(
                    rows,
                    "appointment-present",
                    "Not present",
                    "Present",
                    ReviewStatusCodes.Fail,
                    ReviewSeverityCodes.Error,
                    findingPrefix +
                    "Synthetic appointment metadata is absent.");
                SetCheckFinding(
                    rows,
                    "ct-age-days",
                    "45",
                    "<= 14",
                    ReviewStatusCodes.Variation,
                    ReviewSeverityCodes.Warning,
                    findingPrefix +
                    "Synthetic image age exceeds the configured interval.");
                SetCheckFinding(
                    rows,
                    "user-origin-set",
                    "Not set",
                    "Set",
                    ReviewStatusCodes.Fail,
                    ReviewSeverityCodes.Error,
                    findingPrefix +
                    "Synthetic user origin is absent.");
                SetCheckFinding(
                    rows,
                    "hu-table",
                    "Synthetic alternative table",
                    "Synthetic default table",
                    ReviewStatusCodes.Fail,
                    ReviewSeverityCodes.Error,
                    findingPrefix +
                    "Synthetic calibration table differs from the expected table.");
                SetCheckFinding(
                    rows,
                    "treatment-approval",
                    "Unapproved",
                    "Approved",
                    ReviewStatusCodes.Fail,
                    ReviewSeverityCodes.Error,
                    findingPrefix +
                    "Synthetic treatment approval is absent.");
            }

            return rows;
        }

        private static ReviewCheckRow Check(
            string code,
            string category,
            string observed,
            string expected,
            string unit,
            string message)
        {
            return new ReviewCheckRow
            {
                CheckCode = code,
                Category = category,
                Status = ReviewStatusCodes.Pass,
                Severity = ReviewSeverityCodes.None,
                ObservedValue = observed,
                ExpectedValue = expected,
                Unit = unit,
                Message = message
            };
        }

        private static void SetCheckFinding(
            IEnumerable<ReviewCheckRow> rows,
            string code,
            string observed,
            string expected,
            string status,
            string severity,
            string message)
        {
            ReviewCheckRow row = rows.Single(item => item.CheckCode == code);
            row.ObservedValue = observed;
            row.ExpectedValue = expected;
            row.Status = status;
            row.Severity = severity;
            row.Message = message;
        }

        private static List<ReviewFieldRow> CreateFieldRows(
            bool useFaultExample)
        {
            IList<BeamNamingInput> beams = useFaultExample
                ? CreateFaultFieldInputs()
                : CreateConformingFieldInputs();
            IList<FieldNameSuggestion> suggestions =
                FieldNameSuggester.Suggest("synthetic-plan", beams);

            return suggestions.Select(suggestion =>
                new ReviewFieldRow
                {
                    StableId =
                        "field-" + suggestion.DisplayOrder.ToString("00"),
                    TreatmentOrder = suggestion.DisplayOrder,
                    BeamNumber = suggestion.BeamNumber,
                    CurrentId = suggestion.CurrentId,
                    ExpectedId = suggestion.ExpectedId,
                    CurrentName = useFaultExample
                        ? "Legacy label " +
                          suggestion.DisplayOrder.ToString("00")
                        : suggestion.SuggestedName,
                    SuggestedName = suggestion.SuggestedName,
                    IdStatus = suggestion.IdWouldChange
                        ? ReviewStatusCodes.Fail
                        : ReviewStatusCodes.Pass,
                    NameStatus = useFaultExample
                        ? ReviewStatusCodes.Fail
                        : ReviewStatusCodes.Pass
                }).ToList();
        }

        private static IList<BeamNamingInput> CreateConformingFieldInputs()
        {
            return new[]
            {
                Arc("SYN01", 1, 0, 181.0, 179.0, 0.0, "CW"),
                Arc("SYN02", 2, 1, 179.0, 181.0, 0.0, "CCW"),
                Static("SYN03", 3, 2, 0.0),
                Static("SYN04", 4, 3, 90.0),
                Static("SYN05", 5, 4, 270.0)
            };
        }

        private static IList<BeamNamingInput> CreateFaultFieldInputs()
        {
            return new[]
            {
                Arc("SYN01", 1, 0, 179.0, 181.0, 30.0, "CW"),
                Arc("SYN02", 2, 1, 179.0, 181.0, 0.0, "CW"),
                Arc("SYN03", 3, 2, 179.0, 181.0, 0.0, "CW"),
                Arc("SYN04", 4, 3, 179.0, 181.0, 0.0, "CCW"),
                Arc("SYN05", 5, 4, 179.0, 181.0, 0.0, "CCW")
            };
        }

        private static BeamNamingInput Arc(
            string id,
            int beamNumber,
            int treatmentOrder,
            double startAngle,
            double stopAngle,
            double patientSupportAngle,
            string direction)
        {
            return new BeamNamingInput
            {
                CurrentId = id,
                BeamNumber = beamNumber,
                TreatmentOrderIndex = treatmentOrder,
                GantryStartAngle = startAngle,
                GantryStopAngle = stopAngle,
                PatientSupportAngle = patientSupportAngle,
                GantryDirection = direction
            };
        }

        private static BeamNamingInput Static(
            string id,
            int beamNumber,
            int treatmentOrder,
            double gantryAngle)
        {
            return new BeamNamingInput
            {
                CurrentId = id,
                BeamNumber = beamNumber,
                TreatmentOrderIndex = treatmentOrder,
                GantryStartAngle = gantryAngle,
                GantryStopAngle = gantryAngle,
                PatientSupportAngle = 0.0,
                GantryDirection = string.Empty
            };
        }

        private static List<ReviewStructureMapping> CreateMappings(
            bool useFaultExample)
        {
            var mappings = new List<ReviewStructureMapping>
            {
                Mapping("mapping-target", "Target", "PTV_60"),
                Mapping("mapping-cord", "Spinal cord", "SpinalCord"),
                Mapping("mapping-parotid-left", "Parotid left", "Parotid_L"),
                Mapping("mapping-parotid-right", "Parotid right", "Parotid_R")
            };

            if (useFaultExample)
            {
                ReviewStructureMapping ambiguous = mappings[2];
                ambiguous.TemplateStructure = "Parotid";
                ambiguous.AvailableStructureIds = new List<string>
                {
                    "Parotid_L",
                    "Parotid_R"
                };
                ambiguous.Status = ReviewStatusCodes.Variation;
                ambiguous.Message =
                    "Ambiguous synthetic alias; verify the selected structure.";
            }

            return mappings;
        }

        private static ReviewStructureMapping Mapping(
            string stableId,
            string templateStructure,
            string selectedStructureId)
        {
            return new ReviewStructureMapping
            {
                StableId = stableId,
                TemplateStructure = templateStructure,
                SelectedStructureId = selectedStructureId,
                AvailableStructureIds = new List<string>
                {
                    selectedStructureId
                },
                Status = ReviewStatusCodes.Pass,
                Message = "Synthetic structure mapping conforms."
            };
        }

        private static List<ReviewDvhSeries> CreateDvhSeries(
            string scenarioId)
        {
            double targetMidpointGy =
                scenarioId == "target-underdose" ||
                scenarioId == "mixed-review"
                    ? 58.8
                    : 60.8;
            double cordScaleGy =
                scenarioId == "oar-overdose" ||
                scenarioId == "mixed-review"
                    ? 32.0
                    : 22.0;
            double parotidScaleGy =
                scenarioId == "oar-overdose" ||
                scenarioId == "mixed-review"
                    ? 38.0
                    : 25.0;

            return new List<ReviewDvhSeries>
            {
                Series(
                    "dvh-target",
                    "PTV_60",
                    "PTV 60 Gy",
                    ReviewDvhRoleCodes.Target,
                    "#D1495B",
                    true,
                    182.0,
                    SyntheticDvhFactory.CreateLogisticCumulativePercent(
                        70.0,
                        1.0,
                        targetMidpointGy,
                        0.8)),
                Series(
                    "dvh-cord",
                    "SpinalCord",
                    "Spinal cord",
                    ReviewDvhRoleCodes.OrganAtRisk,
                    "#00798C",
                    true,
                    54.0,
                    SyntheticDvhFactory.CreateExponentialCumulativePercent(
                        70.0,
                        1.0,
                        cordScaleGy,
                        3.0)),
                Series(
                    "dvh-parotid-left",
                    "Parotid_L",
                    "Parotid left",
                    ReviewDvhRoleCodes.OrganAtRisk,
                    "#EDA129",
                    true,
                    27.0,
                    SyntheticDvhFactory.CreateExponentialCumulativePercent(
                        70.0,
                        1.0,
                        parotidScaleGy,
                        1.5)),
                Series(
                    "dvh-parotid-right",
                    "Parotid_R",
                    "Parotid right",
                    ReviewDvhRoleCodes.OrganAtRisk,
                    "#6A4C93",
                    false,
                    26.0,
                    SyntheticDvhFactory.CreateExponentialCumulativePercent(
                        70.0,
                        1.0,
                        23.0,
                        1.5)),
                Series(
                    "dvh-external",
                    "External",
                    "External",
                    ReviewDvhRoleCodes.External,
                    "#64748B",
                    false,
                    15600.0,
                    SyntheticDvhFactory.CreateExponentialCumulativePercent(
                        70.0,
                        1.0,
                        35.0,
                        1.7))
            };
        }

        private static bool UsesFieldAndMappingFindings(string scenarioId)
        {
            return scenarioId == "field-and-mapping" ||
                scenarioId == "mixed-review";
        }

        private static bool UsesOptionalFallback(string scenarioId)
        {
            return scenarioId == "optional-path-fallback" ||
                scenarioId == "mixed-review";
        }

        private static ReviewDvhSeries Series(
            string stableId,
            string structureId,
            string displayName,
            string role,
            string colorHex,
            bool selected,
            double volumeCc,
            List<ReviewDvhPoint> points)
        {
            return new ReviewDvhSeries
            {
                StableId = stableId,
                StructureId = structureId,
                DisplayName = displayName,
                Role = role,
                ColorHex = colorHex,
                LineStyle = ReviewLineStyleCodes.Solid,
                Selected = selected,
                VolumeCc = volumeCc,
                Points = points
            };
        }
    }
}
