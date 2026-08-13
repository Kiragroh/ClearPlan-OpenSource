using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Length >= 2 &&
                args[0].Equals("RefDbContract", StringComparison.OrdinalIgnoreCase))
            {
                return RefDbContractTests.Run(args[1]);
            }

            if (args.Length >= 1 &&
                args[0].Equals(
                    "ExportScenarios",
                    StringComparison.OrdinalIgnoreCase))
            {
                return SyntheticScenarioTests.ExportScenarios(
                    args.Length >= 2 ? args[1] : null);
            }

            if (args.Length >= 2 &&
                args[0].Equals(
                    "ValidateSnapshot",
                    StringComparison.OrdinalIgnoreCase))
            {
                return ValidateExternalSnapshot(args[1]);
            }

            string filter = args.FirstOrDefault() ?? string.Empty;
            var tests = new List<TestCase>
            {
                new TestCase("Smoke.ConstraintSourceMode", SmokeTests.ConstraintSourceModeIsStable),
                new TestCase("ConstraintValueNormalizer.Comparators", ConstraintValueNormalizerTests.NormalizesComparators),
                new TestCase("ConstraintValueNormalizer.Units", ConstraintValueNormalizerTests.NormalizesUnits),
                new TestCase("ConstraintValueNormalizer.Decimals", ConstraintValueNormalizerTests.ParsesInvariantAndGermanDecimals),
                new TestCase("ConstraintValueNormalizer.Metrics", ConstraintValueNormalizerTests.NormalizesSupportedMetrics),
                new TestCase("ConstraintCatalogValidator.Direction", ConstraintCatalogValidatorTests.RejectsComparatorAgainstMetricDirection),
                new TestCase("ConstraintCatalogValidator.Identity", ConstraintCatalogValidatorTests.ReportsDuplicateAndMissingIdentities),
                new TestCase("RefDbJsonConstraintSource.Schema", RefDbJsonConstraintSourceTests.RejectsUnexpectedSchema),
                new TestCase("RefDbJsonConstraintSource.Active", RefDbJsonConstraintSourceTests.LoadsOnlyActiveEntitiesByDefault),
                new TestCase("RefDbJsonConstraintSource.Aliases", RefDbJsonConstraintSourceTests.RetainsAliasesAndSideAliases),
                new TestCase("RefDbJsonConstraintSource.Constraints", RefDbJsonConstraintSourceTests.JoinsAndNormalizesConstraints),
                new TestCase("StructureAliasResolver.Order", StructureAliasResolverTests.UsesDeterministicRuleOrder),
                new TestCase("StructureAliasResolver.Laterality", StructureAliasResolverTests.ProtectsLeftAndRightLaterality),
                new TestCase("StructureAliasResolver.Expansion", StructureAliasResolverTests.ExpandsPairedButNotCombinedStructures),
                new TestCase("StructureAliasResolver.Raw", StructureAliasResolverTests.ResolvesMissingStructureIdFromRawName),
                new TestCase("StructureAliasResolver.Fallback", StructureAliasResolverTests.UsesCodeBeforeDicomTypeFallback),
                new TestCase("StructureAliasResolver.Ambiguity", StructureAliasResolverTests.ReportsAmbiguityInsteadOfChoosingArbitrarily),
                new TestCase("SyntheticDvhMetrics.Interpolation", SyntheticDvhMetricsTests.InterpolatesEndpointsAndMidpoints),
                new TestCase("SyntheticDvhMetrics.MeanDose", SyntheticDvhMetricsTests.IntegratesMeanDose),
                new TestCase("SyntheticDvhMetrics.Thresholds", SyntheticDvhMetricsTests.ClassifiesExactThresholdBoundaries),
                new TestCase("SyntheticDvhMetrics.Invalid", SyntheticDvhMetricsTests.RejectsInvalidCurvesAndRequests),
                new TestCase("ExcelConstraintSource.SharedStrings", ExcelConstraintSourceTests.LoadsSharedStringsAndUnicode),
                new TestCase("ExcelConstraintSource.Schema", ExcelConstraintSourceTests.ReportsMissingHeaders),
                new TestCase("ExcelConstraintSource.Relationships", ExcelConstraintSourceTests.ReportsDuplicateIdsAndBrokenRelationships),
                new TestCase("ConstraintCatalogService.AutomaticRefDb", ConstraintCatalogServiceTests.AutomaticUsesValidRefDbFirst),
                new TestCase("ConstraintCatalogService.AutomaticFallback", ConstraintCatalogServiceTests.AutomaticFallsBackToExcelWithWarning),
                new TestCase("ConstraintCatalogService.Explicit", ConstraintCatalogServiceTests.ExplicitModesNeverSilentlyFallback),
                new TestCase("ConstraintCatalogService.Blocked", ConstraintCatalogServiceTests.AutomaticBlocksWhenNoSourceIsUsable),
                new TestCase("ConstraintCatalogService.Paths", ConstraintCatalogServiceTests.ResolvesRelativeAndAbsolutePaths),
                new TestCase("ConstraintTableSelector.Fractionation", ConstraintTableSelectorTests.ExactFractionationOutranksGenericTables),
                new TestCase("ConstraintTableSelector.Dose", ConstraintTableSelectorTests.DoseRangesRefineCompatibleCandidates),
                new TestCase("ConstraintTableSelector.PlanSum", ConstraintTableSelectorTests.PlanSumRequiresPlanSumTable),
                new TestCase("ConstraintTableSelector.Coverage", ConstraintTableSelectorTests.StructureCoverageCannotRescueIncompatibleFractionation),
                new TestCase("ConstraintTableSelector.Ambiguity", ConstraintTableSelectorTests.TiesRequireConfirmation),
                new TestCase("ConstraintTableSelector.LegacyNine", ConstraintTableSelectorTests.SelectsAllNineLegacyDefaults),
                new TestCase("PqmObjectiveMapper.Contract", PqmObjectiveMappingContractTests.MapsNormalizedConstraintToLegacyPqmContract),
                new TestCase("PqmObjectiveMapper.Fallbacks", PqmObjectiveMappingContractTests.SuppliesSafeIdentityFallbacks),
                new TestCase("PqmObjectiveMapper.Null", PqmObjectiveMappingContractTests.RejectsMissingConstraint),
                new TestCase("FieldNameSuggester.Basic", FieldNamingPreviewTests.NamesStaticAndArcFields),
                new TestCase("FieldNameSuggester.Couch", FieldNamingPreviewTests.InsertsCouchAngleBeforeDirection),
                new TestCase("FieldNameSuggester.CouchDuplicate", FieldNamingPreviewTests.KeepsCouchAnglePenultimateForDuplicateArcs),
                new TestCase("FieldNameSuggester.Duplicates", FieldNamingPreviewTests.AddsAttachedDuplicateSuffixes),
                new TestCase("FieldNameSuggester.Order", FieldNamingPreviewTests.UsesTreatmentOrderThenBeamNumber),
                new TestCase("FieldNameSuggester.Tolerance", FieldNamingPreviewTests.UsesHalfDegreeTolerance),
                new TestCase("FieldNameSuggester.PlanId", FieldNamingPreviewTests.UsesPlanIdForExpectedFieldIds),
                new TestCase("FieldNameSuggester.Prefix", FieldNamingPreviewTests.PreservesConformingPrefixFromCurrentFields),
                new TestCase("FieldNameSuggester.Status", FieldNamingPreviewTests.ReportsIdAndNameConformanceSeparately),
                new TestCase("DvhSelection.Exact", DvhSelectionPolicyTests.SelectsExactMappedStructures),
                new TestCase("DvhSelection.External", DvhSelectionPolicyTests.ExcludesExternalAndBodyContours),
                new TestCase("DvhSelection.Partial", DvhSelectionPolicyTests.ExcludesPartialHelpersButKeepsTargets),
                new TestCase("ClinicalReviewMapper.PlanCheck", ClinicalReviewValueMapperTests.MapsPlanCheckStatusAndSeverity),
                new TestCase("ClinicalReviewMapper.Pqm", ClinicalReviewValueMapperTests.MapsPqmStatusAndSeverity),
                new TestCase("ClinicalReviewMapper.StableIds", ClinicalReviewValueMapperTests.CreatesDeterministicUniqueStableIds),
                new TestCase("ClinicalReviewMapper.Decimals", ClinicalReviewValueMapperTests.ParsesInvariantAndGermanDecimals),
                new TestCase("ClinicalReviewMapper.Dose", ClinicalReviewValueMapperTests.ConvertsCentigrayToGray),
                new TestCase("ClinicalReviewMapper.Fields", ClinicalReviewValueMapperTests.MapsFieldConformance),
                new TestCase("ClinicalReviewMapper.Sanitization", ClinicalReviewValueMapperTests.SanitizesPathsAndDicomUids),
                new TestCase("ReviewWorkspaceHost.ActionsOnce", ReviewWorkspaceHostControllerTests.DispatchesVisibleActionsExactlyOnce),
                new TestCase("ReviewWorkspaceHost.RefreshFallback", ReviewWorkspaceHostControllerTests.KeepsLastGoodViewModelWhenRefreshFails),
                new TestCase("ReviewWorkspaceHost.Dispose", ReviewWorkspaceHostControllerTests.UnsubscribesActionsWhenDisposed),
                new TestCase("ReviewWorkspaceHost.SingularHint", ReviewWorkspaceHostControllerTests.UsesSingularHintLabelForOneVariation),
                new TestCase("ReviewSnapshot.ValidMinimal", ReviewSnapshotTests.AcceptsValidMinimalSnapshot),
                new TestCase("ReviewSnapshot.SchemaVersion", ReviewSnapshotTests.RejectsUnsupportedSchemaVersion),
                new TestCase("ReviewSnapshot.Synthetic", ReviewSnapshotTests.RejectsNonSyntheticSimulatorInput),
                new TestCase("ReviewSnapshot.DuplicateIds", ReviewSnapshotTests.RejectsDuplicateStableIds),
                new TestCase("ReviewSnapshot.ActivePlan", ReviewSnapshotTests.RejectsMissingActivePlan),
                new TestCase("ReviewSnapshot.StatusSeverity", ReviewSnapshotTests.RejectsInvalidStatusAndSeverity),
                new TestCase("ReviewSnapshot.FiniteDose", ReviewSnapshotTests.RejectsNonFiniteDose),
                new TestCase("ReviewSnapshot.DoseOrder", ReviewSnapshotTests.RequiresIncreasingDoseOrder),
                new TestCase("ReviewSnapshot.CumulativeVolume", ReviewSnapshotTests.RejectsIncreasingCumulativeVolume),
                new TestCase("ReviewSnapshot.StructureReference", ReviewSnapshotTests.RejectsMissingReferencedStructure),
                new TestCase("ReviewSnapshot.PublishSafeLabels", ReviewSnapshotTests.RejectsIdentityBearingClinicalContent),
                new TestCase("ReviewSnapshot.EmbeddedPrivateContent", ReviewSnapshotTests.RejectsEmbeddedPrivateContentAcrossAllTextSurfaces),
                new TestCase("ReviewSnapshot.GenericSyntheticText", ReviewSnapshotTests.AllowsGenericSyntheticTextAcrossTheContract),
                new TestCase("ReviewSnapshot.JsonRoundTrip", ReviewSnapshotTests.JsonRoundTripIsDeterministic),
                new TestCase("ReviewReportMapping.Sections", ReviewReportMappingTests.PreservesEveryReviewSectionAndUnit),
                new TestCase("ReviewReportMapping.Watermark", ReviewReportMappingTests.ForcesTheSyntheticSafetyMark),
                new TestCase("ReviewReportMapping.PdfSmoke", ReviewReportMappingTests.WritesASyntheticPdf),
                new TestCase("ReviewReportMapping.PdfContent", ReviewReportMappingTests.PopulatesMappedPdfTableCells),
                new TestCase("ReviewWindowSizePolicy.Matrix", ReviewWindowSizePolicyTests.MatchesTheResponsiveWorkAreaMatrix),
                new TestCase("ReviewWindowSizePolicy.Contract", ReviewWindowSizePolicyTests.ExposesTheWindowSizingContract),
                new TestCase("ReviewWindowSizePolicy.Invalid", ReviewWindowSizePolicyTests.RejectsInvalidWorkAreas),
                new TestCase("SyntheticScenario.Catalog", SyntheticScenarioTests.DeclaresExactCatalogAndExpectations),
                new TestCase("SyntheticScenario.RowsAndStatuses", SyntheticScenarioTests.MatchesExactRowsStatusesAndStructures),
                new TestCase("SyntheticScenario.DvhMonotonicity", SyntheticScenarioTests.ProducesValidMonotoneDvhCurves),
                new TestCase("SyntheticScenario.Determinism", SyntheticScenarioTests.SerializesByteIdentically),
                new TestCase("SyntheticScenario.VisibleFaults", SyntheticScenarioTests.ExposesVisibleDoseAndCheckFaults),
                new TestCase("SyntheticScenario.PqmDvhConsistency", SyntheticScenarioTests.PqmValuesAndStatusesMatchDvhCurves),
                new TestCase("SyntheticScenario.FieldAndMapping", SyntheticScenarioTests.UsesFieldNamerAndShowsMappingDeviation),
                new TestCase("SyntheticScenario.OptionalFallback", SyntheticScenarioTests.UsesNeutralOptionalFallbackOnly),
                new TestCase("SyntheticScenario.MixedReview", SyntheticScenarioTests.IntegratedMixedReviewCombinesVisibleFindings),
                new TestCase("SyntheticScenario.Provenance", SyntheticScenarioTests.ContainsSyntheticProvenanceOnly),
                new TestCase("SyntheticScenario.CheckedInJson", SyntheticScenarioTests.CheckedInJsonMatchesFactory),
                new TestCase("SimulatorArguments.Defaults", SimulatorArgumentsTests.UsesDocumentedDefaults),
                new TestCase("SimulatorArguments.ScenarioTab", SimulatorArgumentsTests.ParsesScenarioAndTab),
                new TestCase("SimulatorArguments.Capture", SimulatorArgumentsTests.ParsesSingleCapture),
                new TestCase("SimulatorArguments.CaptureAll", SimulatorArgumentsTests.ParsesCaptureAll),
                new TestCase("SimulatorArguments.LayoutProbe", SimulatorArgumentsTests.ParsesLayoutProbe),
                new TestCase("SimulatorArguments.LayoutProbeInvalid", SimulatorArgumentsTests.RejectsInvalidLayoutProbeRequests),
                new TestCase("SimulatorArguments.Invalid", SimulatorArgumentsTests.RejectsUnknownOrIncompleteValues),
                new TestCase("SimulatorArguments.Conflicts", SimulatorArgumentsTests.RejectsConflictingOrDuplicateValues),
                new TestCase("SimulatorArguments.OutputPaths", SimulatorArgumentsTests.RejectsUnsafeOrWrongOutputPaths),
                new TestCase("SimulatorOutputPaths.Repository", SimulatorOutputPathsTests.UsesRepositoryArtifactsReportDirectory),
                new TestCase("SimulatorOutputPaths.Portable", SimulatorOutputPathsTests.UsesPortableArtifactsReportDirectory),
                new TestCase("SimulatorOutputPaths.Network", SimulatorOutputPathsTests.RejectsNetworkExecutableDirectory),
                new TestCase("SimulatorOutputPaths.Dvh", SimulatorOutputPathsTests.UsesDeterministicLocalDvhExportPath),
                new TestCase("SimulatorVisibleActions.Events", SimulatorVisibleActionTests.CommandsRaiseEveryVisibleActionEvent),
                new TestCase("SimulatorVisibleActions.Reset", SimulatorVisibleActionTests.ResetRestoresInitialSelectionAndPlotVisibility),
                new TestCase("SimulatorVisibleActions.OpenPlan", SimulatorVisibleActionTests.OpenPlanStatusIdentifiesTheSyntheticActivePlan),
                new TestCase("SimulatorVisibleActions.Mapping", SimulatorVisibleActionTests.MappingPersistsIntoSnapshotAndReport),
                new TestCase("SimulatorVisibleActions.MappingInvalid", SimulatorVisibleActionTests.MappingRejectsUnknownChoicesWithoutMutation),
                new TestCase("SimulatorVisibleActions.DvhPng", SimulatorVisibleActionTests.DvhExportWritesALocalPng),
                new TestCase("SimulatorVisibleActions.DvhNetwork", SimulatorVisibleActionTests.DvhExportRejectsNetworkOutput),
                new TestCase("SettingsValidator.RequiredSource", SettingsValidatorTests.RequiresConfiguredSourceByMode),
                new TestCase("SettingsValidator.Extension", SettingsValidatorTests.ChecksSourceExtensionsAndReadability),
                new TestCase("SettingsValidator.Directories", SettingsValidatorTests.AcceptsWritableOperationalDirectories),
                new TestCase("PathSettingsIni.RoundTrip", PathSettingsIniTests.RoundTripsEveryConfiguredPath),
                new TestCase("PathSettingsIni.Legacy", PathSettingsIniTests.ReadsLegacyPathAliases),
                new TestCase("PathFailureClassifier.Index", PathFailureClassifierTests.RejectsCollectionIndexErrors),
                new TestCase("PathFailureClassifier.WrappedIo", PathFailureClassifierTests.RecognizesWrappedIoErrors),
                new TestCase("PqmDefaultReviewPolicy.Valid", PqmDefaultReviewPolicyTests.KeepsValidRowsVisible),
                new TestCase("PqmDefaultReviewPolicy.Partial", PqmDefaultReviewPolicyTests.IgnoresDelimitedPartialHelpers),
                new TestCase("PqmDefaultReviewPolicy.Unavailable", PqmDefaultReviewPolicyTests.UsesExplicitUnavailableStatusForNonTargets),
                new TestCase("WritablePathFallback.File", WritablePathFallbackTests.UsesLocalFileWhenConfiguredParentIsUnavailable),
                new TestCase("WritablePathFallback.Directory", WritablePathFallbackTests.KeepsConfiguredDirectoryWhenItIsWritable)
            };

            var selected = tests
                .Where(test => string.IsNullOrWhiteSpace(filter) ||
                               test.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            if (selected.Count == 0)
            {
                Console.Error.WriteLine("No tests matched filter '{0}'.", filter);
                return 2;
            }

            int failures = 0;
            foreach (TestCase test in selected)
            {
                try
                {
                    test.Execute();
                    Console.WriteLine("PASS {0}", test.Name);
                }
                catch (Exception exception)
                {
                    failures++;
                    Console.Error.WriteLine("FAIL {0}: {1}", test.Name, exception.Message);
                }
            }

            Console.WriteLine("{0} tests, {1} failures", selected.Count, failures);
            return failures == 0 ? 0 : 1;
        }

        private static int ValidateExternalSnapshot(string path)
        {
            try
            {
                ReviewSnapshot snapshot = ReviewSnapshotJson.DeserializeUtf8(
                    File.ReadAllBytes(path));
                ReviewSnapshotValidationResult result =
                    ReviewSnapshotValidator.Validate(snapshot);
                foreach (ReviewSnapshotValidationIssue issue in result.Issues)
                {
                    Console.Error.WriteLine(
                        "{0} {1}: {2}",
                        issue.Code,
                        issue.Path,
                        issue.Message);
                }

                Console.WriteLine(
                    "External snapshot validation: {0} issue(s).",
                    result.Issues.Count);
                return result.IsValid ? 0 : 1;
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(
                    "External snapshot validation failed: " + error.Message);
                return 2;
            }
        }

        private sealed class TestCase
        {
            public TestCase(string name, Action execute)
            {
                Name = name;
                Execute = execute;
            }

            public string Name { get; private set; }
            public Action Execute { get; private set; }
        }
    }
}
