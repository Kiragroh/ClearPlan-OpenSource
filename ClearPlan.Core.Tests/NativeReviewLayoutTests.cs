using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Presentation.Plot;
using ClearPlan.Presentation.ViewModels;
using OxyPlot;
using OxyPlot.Series;

namespace ClearPlan.Core.Tests
{
    internal static class NativeReviewLayoutTests
    {
        public static void ParameterPlotsKeepNominalValueWithoutGeometry()
        {
            var beam = new ReviewBeamAnalysis { BeamId = "SYNTHETIC", NominalDoseRateMuPerMin = 600 };
            var vm = new PlanAnalysisViewModel(new ReviewPlanAnalysis { Beams = new List<ReviewBeamAnalysis> { beam } });
            var property = vm.GetType().GetProperty("NominalDoseRateText");
            TestAssert.NotNull(property, "The nominal field setting must remain visible when control-point extraction is unavailable.");
            string label = (string)property.GetValue(vm, null);
            TestAssert.True(label.Contains("600 MU/min") && label.Contains("Nominal"), "Label the field value as nominal, not delivered.");
            TestAssert.Equal(0, vm.DoseRatePlotModel.Series.Count, "Do not invent control-point indices when only the beam setting exists.");
        }

        public static void ParameterPlotsUseRealIndicesAndVisibleMarkers()
        {
            var beam = new ReviewBeamAnalysis { NominalDoseRateMuPerMin = 600 };
            beam.ControlPoints.Add(new ReviewControlPointSample { Index = 7, ApertureAreaCm2 = 12 });
            beam.ControlPoints.Add(new ReviewControlPointSample { Index = 9 });
            beam.ControlPoints.Add(new ReviewControlPointSample { Index = 11, ApertureAreaCm2 = 18 });
            var vm = new PlanAnalysisViewModel(new ReviewPlanAnalysis { Beams = new List<ReviewBeamAnalysis> { beam } });
            TestAssert.Equal(0, vm.DoseRatePlotModel.Series.Count, "A nominal field setting is not a control-point rate trajectory, even when control points exist.");
            var area = (LineSeries)vm.AperturePlotModel.Series[0];
            TestAssert.True(double.IsNaN(area.Points[1].Y), "Missing aperture samples must remain gaps, not zero or interpolated geometry.");
            TestAssert.True(area.MarkerType != MarkerType.None && area.MarkerSize >= 2,
                "A single sample or samples separated by gaps need visible markers, not an invisible one-point line.");
            TestAssert.False(vm.DoseRatePlotModel.Series.Any(s => s.Title.Contains("Geplant")), "Never fabricate planned rates from nominal values.");
        }

        public static void ParameterPlotsExposeUnavailableAndChangedBeamState()
        {
            var unavailable = new ReviewBeamAnalysis { ControlPoints = null };
            var vm = new PlanAnalysisViewModel(new ReviewPlanAnalysis { Beams = new List<ReviewBeamAnalysis> { unavailable } });
            TestAssert.Equal(0, vm.DoseRatePlotModel.Series.Count);
            TestAssert.Equal(0, vm.AperturePlotModel.Series.Count);
            var changed = new List<string>();
            vm.PropertyChanged += (sender, args) => changed.Add(args.PropertyName);
            vm.SelectedBeam = new ReviewBeamAnalysis { NominalDoseRateMuPerMin = 800 };
            TestAssert.True(changed.Contains("NominalDoseRateText"), "Changing beam must refresh the visible nominal setting.");
            string xaml = File.ReadAllText(Path.Combine("ClearPlan.Presentation", "Views", "PlanParametersView.xaml"));
            TestAssert.True(xaml.Contains("{lang:Translate Path=NominalDoseRateText}"), "Bind the nominal readout beside the dose-rate plot.");
            TestAssert.True(xaml.Contains("Click=\"ShowParameterPlots\""), "Keep below-fold plots directly reachable from the section header.");
        }

        public static void ParameterRatePlotsUseOnlyAvailablePlannedSamples()
        {
            var beam = new ReviewBeamAnalysis { NominalDoseRateMuPerMin = 600 };
            beam.ControlPoints.Add(new ReviewControlPointSample { Index = 0, PlannedDoseRateMuPerMin = 200 });
            beam.ControlPoints.Add(new ReviewControlPointSample { Index = 1, NominalDoseRateMuPerMin = 600 });
            beam.ControlPoints.Add(new ReviewControlPointSample { Index = 2, PlannedDoseRateMuPerMin = 500 });
            beam.ControlPoints.Add(new ReviewControlPointSample { Index = 3, PlannedDoseRateMuPerMin = -1 });
            beam.ControlPoints.Add(new ReviewControlPointSample { Index = 4, PlannedDoseRateMuPerMin = double.PositiveInfinity });
            var vm = new PlanAnalysisViewModel(new ReviewPlanAnalysis { Beams = new List<ReviewBeamAnalysis> { beam } });
            TestAssert.Equal(1, vm.DoseRatePlotModel.Series.Count, "Only the supplied planned rate belongs in the trajectory plot.");
            var rate = (LineSeries)vm.DoseRatePlotModel.Series[0];
            TestAssert.True(rate.Title.Contains("Geplant"));
            TestAssert.True(rate.Points.Select(p => p.X).SequenceEqual(new[] { 0d, 1d, 2d, 3d, 4d }));
            TestAssert.True(rate.Points[0].Y == 200 && rate.Points[2].Y == 500);
            TestAssert.True(double.IsNaN(rate.Points[1].Y) && double.IsNaN(rate.Points[3].Y) && double.IsNaN(rate.Points[4].Y),
                "Missing or invalid rates remain gaps; never substitute the nominal setting.");
            var hasRate = vm.GetType().GetProperty("HasPlannedDoseRate");
            TestAssert.NotNull(hasRate, "Expose an explicit rate-availability state for the view.");
            TestAssert.True((bool)hasRate.GetValue(vm));
            var changed = new List<string>();
            vm.PropertyChanged += (s, e) => changed.Add(e.PropertyName);
            vm.SelectedBeam = new ReviewBeamAnalysis { NominalDoseRateMuPerMin = 800 };
            TestAssert.False((bool)hasRate.GetValue(vm));
            TestAssert.True(changed.Contains("HasPlannedDoseRate"));
            string xaml = File.ReadAllText(Path.Combine("ClearPlan.Presentation", "Views", "PlanParametersView.xaml"));
            TestAssert.True(xaml.Contains("DoseRateUnavailable") && xaml.Contains("HasDoseRateTrace"));
            var estimatedBeam = new ReviewBeamAnalysis { DoseRateEstimateStatus = "Estimated", DoseRateEstimateProfile = "Test profile" };
            estimatedBeam.ControlPoints.Add(new ReviewControlPointSample { Index = 0 });
            estimatedBeam.ControlPoints.Add(new ReviewControlPointSample { Index = 1, EstimatedDoseRateMuPerMin = 150 });
            estimatedBeam.ControlPoints.Add(new ReviewControlPointSample { Index = 2, EstimatedDoseRateMuPerMin = 600 });
            vm.SelectedBeam = estimatedBeam;
            TestAssert.False(vm.HasPlannedDoseRate, "An estimate is not an ESAPI-supplied planned rate.");
            TestAssert.Equal(1, vm.DoseRatePlotModel.Series.Count, "A valid estimate must be visible without planned rate samples.");
            var estimate = (LineSeries)vm.DoseRatePlotModel.Series[0];
            TestAssert.True(estimate.Title.Contains("Geschätzt") && vm.DoseRatePlotModel.Subtitle.Contains("keine gemessene"));
            TestAssert.True(double.IsNaN(estimate.Points[0].Y) && estimate.Points[1].Y == 150 && estimate.Points[2].Y == 600);
            TestAssert.True(estimate.LineStyle == LineStyle.Dash, "Estimates must be visually distinct from supplied planned samples.");
            estimatedBeam.DoseRateEstimateStatus = "Unavailable";
            vm.SelectedBeam = estimatedBeam;
            TestAssert.Equal(0, vm.DoseRatePlotModel.Series.Count, "Do not draw stale estimates from an unavailable calculation.");
        }

        public static void ReciprocalConformityUsesPaddickNaming()
        {
            string xaml = File.ReadAllText(Path.Combine("ClearPlan.Presentation", "Views", "PlanParametersView.xaml"));
            TestAssert.True(xaml.Contains("Header=\"{lang:Translate Value='1 / Paddick CI'}\""), "Name reciprocal conformity explicitly in the GUI.");
            TestAssert.False(xaml.Contains("CI (PlanCheck)"));
            TestAssert.False(TargetQualityCalculator.Definition.Contains("CI (PlanCheck)"));
            TestAssert.True(TargetQualityCalculator.Definition.Contains("1 / Paddick CI = (TV × V_Rx) / TV_Rx²"));
            var analysis = new ReviewPlanAnalysis();
            analysis.TargetQuality.Add(TargetQualityCalculator.Calculate("SYNTHETIC_PTV", 60, 100, 90, 120, 360, 63, 57));
            var document = new MigraDoc.DocumentObjectModel.Document();
            var section = document.AddSection();
            var method = typeof(ClearPlan.Reporting.MigraDoc.ReportPdf).GetMethod("AddTargetQuality",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            method.Invoke(new ClearPlan.Reporting.MigraDoc.ReportPdf(), new object[] { section, analysis });
            string ddl = MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToString(document);
            TestAssert.True(ddl.Contains("1 / Paddick CI") && !ddl.Contains("CI (PlanCheck)"), "PDF and GUI must use the same reciprocal label.");
            TestAssert.True(Math.Abs(analysis.TargetQuality[0].PaddickCi.Value * analysis.TargetQuality[0].PlanCheckCi.Value - 1) < 1e-12,
                "This wording change must preserve reciprocal CI calculations.");
        }

        public static void HeaderFitsClinicalAndSandboxContent()
        {
            string xaml = File.ReadAllText(Path.Combine("ClearPlan.Script", "MainView.xaml"));
            int body = xaml.IndexOf("<Grid Grid.Row=\"1\">", StringComparison.Ordinal);
            string header = xaml.Substring(0, body);
            TestAssert.False(header.Contains("<RowDefinition Height=\"70\"/>"),
                "The native header must measure all clinical and sandbox lines, not clip at 70 pixels.");
            TestAssert.True(header.Contains("x:Name=\"NativeHeader\"") &&
                header.Contains("x:Name=\"HeaderConstraintRow\""),
                "Keep title/actions and constraint controls in separately measured header rows.");
            TestAssert.True(header.Contains("x:Name=\"ConfirmConstraintButton\"") &&
                header.Contains("MinHeight=\"36\""), "Constraint confirmation must have a full-height target.");
            TestAssert.True(header.Contains("TextWrapping=\"Wrap\""), "Long header status must wrap.");
            TestAssert.True(xaml.Contains("DataContext=\"{x:Null}\""),
                "The shared workspace must not inherit a legacy OxyPlot model during construction.");
        }

        public static void DvhLegendIsOutsideOnTheRight()
        {
            var model = ReviewPlotFactory.Create(null);
            TestAssert.Equal(LegendPlacement.Outside, model.LegendPlacement,
                "The DVH legend must never cover dose-volume curves.");
            TestAssert.Equal(LegendPosition.RightTop, model.LegendPosition);
            TestAssert.True(model.LegendMaxWidth > 0 && model.LegendMaxWidth <= 260,
                "Bound legend width so longer structure names do not consume the plot.");
            string legacy = File.ReadAllText(Path.Combine("ClearPlan.Script", "ViewModels", "MainViewModel.cs"));
            TestAssert.True(legacy.Contains("pm.LegendPlacement = LegendPlacement.Outside;"),
                "The compatibility DVH must use the same outside-right layout.");
        }

        public static void NativePlanSwitchUsesTheExistingReadOnlyReplacement()
        {
            string xaml = File.ReadAllText(Path.Combine("ClearPlan.Script", "MainView.xaml"));
            TestAssert.True(xaml.Contains("x:Name=\"SwitchPlanButton\"") &&
                xaml.Contains("Content=\"{lang:Translate Value='Plan wechseln'}\"") && xaml.Contains("Click=\"SwitchPlanClicked\""),
                "Expose a native plan-switch action without requiring a hidden table action.");
            string source = File.ReadAllText(Path.Combine("ClearPlan.Script", "MainView.xaml.cs"));
            int start = source.IndexOf("private void SwitchPlanClicked(", StringComparison.Ordinal);
            int end = source.IndexOf("private void PlanOpen_Button_Click(", StringComparison.Ordinal);
            TestAssert.True(start >= 0 && end > start, "The plan switch must be wired to a handler.");
            string handler = source.Substring(start, end - start);
            TestAssert.True(handler.Contains("IsSyntheticDemo") && handler.Contains("OpenPlanningItem("),
                "Guard synthetic mode and reuse native plan replacement, not mutable plan state.");
            TestAssert.True(source.Contains("var mainViewModel = new MainViewModel(") &&
                source.Contains("var replacement = new MainView(mainViewModel, comparisonReference);") &&
                source.Contains("_clinicalReviewHost.Dispose();"),
                "Plan changes must create fresh review/mapping state and detach the old host.");
            TestAssert.False(handler.Contains("BeginModifications") || handler.Contains("SaveModifications"));
        }
    }
}
