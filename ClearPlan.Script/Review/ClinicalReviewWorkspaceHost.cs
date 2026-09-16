using System;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.Win32;
using VMS.TPS.Common.Model.API;
using ClearPlan.Dicom;
using ClearPlan.Helpers;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;
using ClearPlan.Presentation.Views;

namespace ClearPlan.Review
{
    /// <summary>
    /// Adapts the detached review workspace to the live, read-only Eclipse
    /// view model. Vendor access stays on MainView's WPF/ESAPI thread; file
    /// parsing and detached target projection may be awaited off that thread.
    /// </summary>
    internal sealed partial class ClinicalReviewWorkspaceHost : IDisposable
    {
        private readonly MainView owner;
        private readonly MainViewModel source;
        private readonly Func<ClearPlanSettings> settingsProvider;
        private readonly ReviewWorkspaceView view;
        private readonly ReviewWorkspaceHostController controller;
        private bool disposed;
        private ParsedRtPlan importedPlan;
        private string importedPlanUid;
        private ClearPlan.Core.PlanAnalysis.NativeMlcProfileCatalog nativeProfiles;
        private ClearPlan.Core.PlanAnalysis.DoseRateEstimationProfileCatalog doseRateProfiles;
        private CancellationTokenSource analysisCancellation;
        internal Task NativeAnalysisTask { get; private set; }
        internal Task PlanImagesTask { get; private set; }
        private Task<System.Collections.Generic.List<ClearPlan.Core.Review.ReviewPlanImage>> imageCapture;
        private ClearPlan.Core.Review.ReviewSnapshot imageCaptureSnapshot;
        private CancellationTokenSource imageCancellation;
        private string explicitTargetOverride;
        internal Task ComparisonTask { get; private set; }
        internal Task HtmlReportTask { get; private set; }
        internal Task AriaUploadTask { get; private set; }
        private readonly ClearPlan.Core.Review.ReviewPlanContextGuard planContext = new ClearPlan.Core.Review.ReviewPlanContextGuard();

        public ClinicalReviewWorkspaceHost(
            MainView ownerValue,
            MainViewModel sourceValue,
            Func<ClearPlanSettings> settingsProviderValue,
            ReviewWorkspaceView viewValue,
            ClearPlan.Core.Review.ReviewSnapshot comparisonReference = null)
        {
            owner = ownerValue ??
                throw new ArgumentNullException("ownerValue");
            source = sourceValue ??
                throw new ArgumentNullException("sourceValue");
            settingsProvider = settingsProviderValue ??
                throw new ArgumentNullException("settingsProviderValue");
            view = viewValue ??
                throw new ArgumentNullException("viewValue");

            var bindings = new ReviewWorkspaceActionBindings
            {
                Report = OnReportRequested,
                HtmlReport = OnHtmlReportRequested,
                AriaUpload = OnAriaUploadRequested,
                OpenPlan = OnOpenPlanRequested,
                SelectComparisonPlan = OnComparisonPlanRequested,
                ResetDvh = OnDvhResetRequested,
                ExportDvh = OnDvhExportRequested,
                Navigate = OnNavigationRequested,
                ApplyStructureMapping = OnStructureMappingRequested,
                SelectConstraintTable = OnConstraintTableRequested,
                OpenSettings = OnSettingsRequested,
                SelectScenario = OnScenarioSelectionRequested,
                ImportPlanData = OnNativePlanDataRequested,
                CalculateTarget = OnCalculateTargetRequested,
                GenerateDrr = OnGenerateDrrRequested,
                PlanImages = OnPlanImagesRequested,
                Collision = OnCollisionRequested
            };
            controller = new ReviewWorkspaceHostController(
                BuildSnapshot,
                bindings,
                comparisonReference);
            view.DataContextChanged += OnWorkspaceDataContextChanged;
        }

        public ReviewWorkspaceViewModel CurrentViewModel
        {
            get { return controller.CurrentViewModel; }
        }

        public bool IsSyntheticDemo
        {
            get
            {
                return controller.CurrentViewModel != null &&
                    controller.CurrentViewModel.IsSynthetic;
            }
        }

        public bool TryRefresh()
        {
            ThrowIfDisposed();
            if (IsSyntheticDemo)
            {
                return true;
            }

            Exception failure;
            if (!controller.TryRefresh(out failure))
            {
                owner.ShowSharedWorkspaceFailure(
                    failure,
                    controller.CurrentViewModel != null);
                return false;
            }

            view.DataContext = controller.CurrentViewModel;
            planContext.Commit(controller.CurrentSnapshot, source.ActivePlanningItem.PlanningItemUID, source.ActivePlanningItem.PlanningItemObject);
            owner.ShowSharedWorkspace();
            QueueNativeAnalysis();
            QueueImagesIfVisible();
            return true;
        }

        public bool TrySetSyntheticDemo(bool enabled)
        {
            ThrowIfDisposed();
            if (collisionCancellation != null) collisionCancellation.Cancel();
            if (imageCancellation != null) imageCancellation.Cancel();
            Exception failure;
            bool succeeded = enabled
                ? controller.TryShowSnapshot(
                    SyntheticScenarioFactory.Create("mixed-review"),
                    out failure)
                : controller.TryRefresh(out failure);
            if (!succeeded)
            {
                owner.ShowSharedWorkspaceFailure(
                    failure,
                    controller.CurrentViewModel != null);
                return false;
            }

            view.DataContext = controller.CurrentViewModel;
            if (enabled) planContext.Clear();
            else planContext.Commit(controller.CurrentSnapshot, source.ActivePlanningItem.PlanningItemUID, source.ActivePlanningItem.PlanningItemObject);
            owner.ShowSharedWorkspace();
            if (!enabled) QueueNativeAnalysis();
            else if (analysisCancellation != null) analysisCancellation.Cancel();
            QueueImagesIfVisible();
            return true;
        }

        public void Dispose()
        {
            if (disposed)
            {
                return;
            }

            disposed = true;
            if (collisionCancellation != null) collisionCancellation.Cancel();
            owner.CancelAriaPreparation();
            view.DataContextChanged -= OnWorkspaceDataContextChanged;
            if (imageCancellation != null) imageCancellation.Cancel();
            if (analysisCancellation != null) analysisCancellation.Cancel();
            importedPlan = null;
            planContext.Clear();
            controller.Dispose();
            view.DataContext = null;
        }

        private ClearPlan.Core.Review.ReviewSnapshot BuildSnapshot()
        {
            if (collisionCancellation != null) collisionCancellation.Cancel();
            if (imageCancellation != null) imageCancellation.Cancel();
            ClearPlanSettings settings = settingsProvider();
            if (settings == null)
            {
                throw new InvalidOperationException(
                    "ClearPlan settings are unavailable.");
            }

            // A refresh deliberately invalidates extended analysis rather than retaining
            // a matched RTPLAN/target projection after potentially changed TPS data.
            importedPlan = null;
            importedPlanUid = null;
            return new EsapiReviewSnapshotBuilder().Build(
                source,
                settings);
        }

        private async void OnImportPlanDataRequested(object sender, ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (IsSyntheticDemo || disposed || analysisCancellation != null) return;
            var plan = source.ActivePlanningItem.PlanningItemObject as PlanSetup;
            if (plan == null) { owner.ShowAnalysisStatus("RTPLAN-Geometrie benötigt einen Einzelplan.", MainView.AnalysisStatusSeverity.Warning); return; }
            var dialog = new OpenFileDialog
            {
                Title = "Zum aktiven Plan passenden DICOM-RTPLAN auswählen",
                Filter = "DICOM-RTPLAN (*.dcm)|*.dcm|Alle Dateien (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false
            };
            if (dialog.ShowDialog() != true) return;
            if (System.Windows.MessageBox.Show(
                "Ist diese Datei ein frischer Export des aktuell geöffneten Plans?\n\n" +
                "SOP-UID, MU und Kontrollpunkte werden geprüft. Eine unveränderte UID allein " +
                "belegt aber nicht, dass nach dem Export keine MLC-Änderungen erfolgt sind. " +
                "Insbesondere beide Halcyon-Lagen sind über diese ESAPI-Version nicht unabhängig abgleichbar.",
                "Aktualität des RTPLAN-Exports bestätigen",
                System.Windows.MessageBoxButton.YesNo,
                System.Windows.MessageBoxImage.Information,
                System.Windows.MessageBoxResult.No) != System.Windows.MessageBoxResult.Yes) return;
            string expectedUid = plan.UID;
            var original = controller.CurrentSnapshot;
            if (!CanExportCurrentSnapshot(original)) { owner.ShowAnalysisStatus("Plankontext ist veraltet; bitte Ansicht aktualisieren.", MainView.AnalysisStatusSeverity.Warning); return; }
            analysisCancellation = new CancellationTokenSource(TimeSpan.FromSeconds(45));
            view.IsEnabled = false;
            owner.ShowAnalysisStatus("RTPLAN wird gelesen und mit dem aktiven Plan abgeglichen …", MainView.AnalysisStatusSeverity.Info);
            try
            {
                // File I/O and DICOM parsing never run on the ESAPI/WPF thread.
                string selectedPath = dialog.FileName;
                var readTask = Task.Run(() => RtPlanReader.Read(selectedPath), analysisCancellation.Token);
                var completed = await Task.WhenAny(readTask, Task.Delay(-1, analysisCancellation.Token));
                if (completed != readTask)
                {
                    // A blocked network read cannot be interrupted safely. Observe a late
                    // failure, but do not allow the result to update this or a later plan.
                    var observeLateFailure = readTask.ContinueWith(task => { var ignored = task.Exception; },
                        TaskContinuationOptions.OnlyOnFaulted);
                    analysisCancellation.Token.ThrowIfCancellationRequested();
                }
                var parsed = await readTask;
                if (!AnalysisContextIsCurrent(original, expectedUid)) return;
                var analysis = RtPlanReader.MatchAndApply(original.PlanAnalysis, parsed, expectedUid);
                owner.ShowAnalysisStatus("MLC-Lagen abgeglichen; Zielprojektionen und PAM werden berechnet …", MainView.AnalysisStatusSeverity.Info);
                analysis = await new EsapiPlanAnalysisBuilder().EnrichTargetProjectionsAsync(plan, analysis, analysisCancellation.Token);
                analysisCancellation.Token.ThrowIfCancellationRequested();
                if (!AnalysisContextIsCurrent(original, expectedUid)) return;
                var updated = ClearPlan.Core.Review.ReviewSnapshotJson.Deserialize(
                    ClearPlan.Core.Review.ReviewSnapshotJson.Serialize(original));
                updated.PlanAnalysis = analysis;
                Exception failure;
                if (!controller.TryShowSnapshot(updated, out failure)) throw failure;
                planContext.Commit(updated, expectedUid, source.ActivePlanningItem.PlanningItemObject);
                importedPlan = parsed;
                importedPlanUid = expectedUid;
                view.DataContext = controller.CurrentViewModel;
                owner.ShowCompletedAnalysis("RTPLAN-Geometrie übernommen.", analysis,
                    "Parameter, Kurven und Report verwenden den abgeglichenen Stand; nicht verfügbare PAM-Werte sind in den Planparametern begründet.");
            }
            catch (OperationCanceledException) { owner.ShowAnalysisStatus("Parameterberechnung abgebrochen / Zeitlimit erreicht; letzter gültiger Stand bleibt erhalten.", MainView.AnalysisStatusSeverity.Warning); }
            catch (Exception)
            {
                // Parser/vendor exception text may contain patient paths or identifiers.
                owner.ShowAnalysisStatus("RTPLAN konnte nicht übernommen werden. Einzelplan, SOP-UID, Fraktionsgruppe und Feld-/Kontrollpunktgeometrie müssen übereinstimmen. Letzter gültiger Stand bleibt erhalten.", MainView.AnalysisStatusSeverity.Error);
            }
            finally
            {
                analysisCancellation.Dispose(); analysisCancellation = null;
                if (!disposed) view.IsEnabled = true;
            }
        }

        private void QueueNativeAnalysis()
        {
            var expectedSnapshot = controller.CurrentSnapshot;
            var previous = NativeAnalysisTask ?? Task.FromResult(0);
            if (analysisCancellation != null) analysisCancellation.Cancel();
            NativeAnalysisTask = StartNativeAnalysisAsync(previous, expectedSnapshot);
        }

        private async Task StartNativeAnalysisAsync(Task previous, ClearPlan.Core.Review.ReviewSnapshot expectedSnapshot)
        {
            await previous;
            // Let the initial workspace render. File I/O and detached projections await
            // workers; every subsequent native read resumes on this owner dispatcher.
            await System.Windows.Threading.Dispatcher.Yield(System.Windows.Threading.DispatcherPriority.Background);
            while (!disposed && ReferenceEquals(controller.CurrentSnapshot, expectedSnapshot) && analysisCancellation != null)
                await Task.Delay(50);
            if (disposed || IsSyntheticDemo || !ReferenceEquals(controller.CurrentSnapshot, expectedSnapshot)) return;
            await RefreshNativeAnalysisAsync();
        }

        private void OnNativePlanDataRequested(object sender, ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (IsSyntheticDemo || disposed || analysisCancellation != null || imageCancellation != null) return;
            QueueNativeAnalysis();
        }

        private async Task RefreshNativeAnalysisAsync()
        {
            if (IsSyntheticDemo || disposed || analysisCancellation != null) return;
            var plan = source.ActivePlanningItem.PlanningItemObject as PlanSetup;
            var original = controller.CurrentSnapshot;
            if (plan == null || !CanExportCurrentSnapshot(original)) { owner.ShowAnalysisStatus("Ein aktueller ESAPI-Einzelplan wird benötigt.", MainView.AnalysisStatusSeverity.Warning); return; }
            string expectedUid = plan.UID;
            analysisCancellation = new CancellationTokenSource(TimeSpan.FromSeconds(45));
            owner.ShowAnalysisStatus("ESAPI-Feldgeometrie wird gelesen; Maschinenprofil aus settings.ini wird geprüft …", MainView.AnalysisStatusSeverity.Info);
            try
            {
                var settings = settingsProvider();
                var profiles = await ClearPlan.Core.PlanAnalysis.NativeMlcProfileFileSource.LoadAsync(
                    settings.ResolvePath(settings.Paths.MlcGeometryProfilesJsonPath), analysisCancellation.Token);
                ClearPlan.Core.PlanAnalysis.DoseRateEstimationProfileCatalog rateProfiles = null;
                bool rateProfileUnavailable = false;
                try
                {
                    rateProfiles = await ClearPlan.Core.PlanAnalysis.DoseRateProfileFileSource.LoadAsync(
                        string.IsNullOrWhiteSpace(settings.Paths.DoseRateProfilesJsonPath) ? null : settings.ResolvePath(settings.Paths.DoseRateProfilesJsonPath), analysisCancellation.Token);
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception) { rateProfileUnavailable = true; }
                analysisCancellation.Token.ThrowIfCancellationRequested();
                if (!AnalysisContextIsCurrent(original, expectedUid)) return;
                var builder = new EsapiPlanAnalysisBuilder(profiles, rateProfiles);
                var analysis = builder.Build(plan, explicitTargetOverride);
                if (rateProfileUnavailable) analysis.Warnings.Add("Dose-rate estimate unavailable: check the JSON profile path in settings. Other plan metrics remain independent.");
                if (analysis.Beams.Any(b => b.ControlPoints.Any(cp => cp.Aperture != null)))
                    analysis = await builder.EnrichTargetProjectionsAsync(plan, analysis, analysisCancellation.Token);
                analysisCancellation.Token.ThrowIfCancellationRequested();
                if (!AnalysisContextIsCurrent(original, expectedUid)) return;
                var updated = ClearPlan.Core.Review.ReviewSnapshotJson.Deserialize(ClearPlan.Core.Review.ReviewSnapshotJson.Serialize(original));
                updated.PlanAnalysis = analysis;
                // Users may select curves while the detached PAM calculation runs.
                // Preserve that current choice instead of restoring the startup defaults.
                var selectedCurves = CurrentViewModel.DvhSeries.ToDictionary(s => s.StableId, s => s.IsSelected);
                var previousAnalysis = CurrentViewModel.Analysis;
                Exception failure;
                if (!controller.TryShowSnapshot(updated, out failure)) throw failure;
                foreach (var series in CurrentViewModel.DvhSeries)
                {
                    bool selected;
                    if (selectedCurves.TryGetValue(series.StableId, out selected)) series.IsSelected = selected;
                }
                RestoreAnalysisSelection(previousAnalysis, CurrentViewModel.Analysis);
                planContext.Commit(updated, expectedUid, source.ActivePlanningItem.PlanningItemObject);
                nativeProfiles = profiles; doseRateProfiles = rateProfiles; importedPlan = null; importedPlanUid = null;
                view.DataContext = CurrentViewModel;
                owner.ShowCompletedAnalysis("Native ESAPI-Parameter aktualisiert.", analysis,
                    "Fehlende Modell-/Lagenzuordnungen sind in den Planparametern mit der beobachteten Arraygröße gekennzeichnet; kein DICOM-Export erforderlich.");
            }
            catch (OperationCanceledException) { if (CanUpdateAnalysisStatus(original)) owner.ShowAnalysisStatus("Native Analyse abgebrochen / Zeitlimit erreicht; letzter gültiger Stand bleibt erhalten.", MainView.AnalysisStatusSeverity.Warning); }
            catch (Exception) { if (CanUpdateAnalysisStatus(original)) owner.ShowAnalysisStatus("Native Analyse nicht übernommen. MLC-Profilpfad in Einstellungen, JSON-Gültigkeit und ESAPI-Plankontext prüfen; letzter gültiger Stand bleibt erhalten.", MainView.AnalysisStatusSeverity.Error); }
            finally { analysisCancellation.Dispose(); analysisCancellation = null; }
        }

        private bool CanUpdateAnalysisStatus(ClearPlan.Core.Review.ReviewSnapshot original)
        {
            return !disposed && !IsSyntheticDemo && ReferenceEquals(controller.CurrentSnapshot, original);
        }

        private static void RestoreAnalysisSelection(PlanAnalysisViewModel previous, PlanAnalysisViewModel next)
        {
            // The displayed target follows the latest calculation, including automatic
            // multi-PTV selection; only beam/control-point browsing is restored here.
            var selectedBeamNumber = previous.SelectedBeam == null ? null : previous.SelectedBeam.BeamNumber;
            int? selectedControlPointIndex = previous.SelectedControlPoint == null ? (int?)null : previous.SelectedControlPoint.Index;
            var beam = next.Beams.FirstOrDefault(b => selectedBeamNumber.HasValue && b.BeamNumber == selectedBeamNumber);
            if (beam == null) return;
            next.SelectedBeam = beam;
            int position = beam.ControlPoints.FindIndex(cp => cp.Index == selectedControlPointIndex);
            if (position >= 0) next.ControlPointPosition = position;
        }

        private async void OnCalculateTargetRequested(object sender, ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (IsSyntheticDemo || disposed || analysisCancellation != null || imageCancellation != null) return;
            var plan = source.ActivePlanningItem.PlanningItemObject as PlanSetup;
            if (plan == null) return;
            string expectedUid = plan.UID;
            var original = controller.CurrentSnapshot;
            if (!CanExportCurrentSnapshot(original)) { owner.ShowAnalysisStatus("Plankontext ist veraltet; bitte Ansicht aktualisieren.", MainView.AnalysisStatusSeverity.Warning); return; }
            analysisCancellation = new CancellationTokenSource(TimeSpan.FromSeconds(45));
            view.IsEnabled = false;
            owner.ShowAnalysisStatus("Zielstruktur wird neu projiziert; der Plan wird nicht verändert …", MainView.AnalysisStatusSeverity.Info);
            try
            {
                var builder = new EsapiPlanAnalysisBuilder(nativeProfiles, doseRateProfiles);
                var analysis = builder.Build(plan, eventArgs.Parameter as string);
                if (importedPlan != null && importedPlanUid == expectedUid)
                    analysis = RtPlanReader.MatchAndApply(analysis, importedPlan, expectedUid);
                analysis = await builder.EnrichTargetProjectionsAsync(plan, analysis, analysisCancellation.Token);
                analysisCancellation.Token.ThrowIfCancellationRequested();
                if (!AnalysisContextIsCurrent(original, expectedUid)) return;
                var updated = ClearPlan.Core.Review.ReviewSnapshotJson.Deserialize(
                    ClearPlan.Core.Review.ReviewSnapshotJson.Serialize(original));
                updated.PlanAnalysis = analysis;
                Exception failure;
                if (!controller.TryShowSnapshot(updated, out failure)) throw failure;
                explicitTargetOverride = eventArgs.Parameter as string;
                planContext.Commit(updated, expectedUid, source.ActivePlanningItem.PlanningItemObject);
                view.DataContext = controller.CurrentViewModel;
                owner.ShowCompletedAnalysis("Planparameter für die gewählte Zielstruktur aktualisiert.", analysis,
                    "Einschränkungen stehen im Methodenhinweis der Planparameter.");
            }
            catch (OperationCanceledException) { owner.ShowAnalysisStatus("PAM-Berechnung abgebrochen / Zeitlimit erreicht; letzter gültiger Stand bleibt erhalten.", MainView.AnalysisStatusSeverity.Warning); }
            catch (Exception) { owner.ShowAnalysisStatus("PAM-Berechnung nicht möglich. Zielstruktur, native ESAPI-Geometrie und MLC-Profil prüfen; letzter gültiger Stand bleibt erhalten.", MainView.AnalysisStatusSeverity.Error); }
            finally
            {
                analysisCancellation.Dispose(); analysisCancellation = null;
                if (!disposed) view.IsEnabled = true;
            }
        }

        private bool AnalysisContextIsCurrent(ClearPlan.Core.Review.ReviewSnapshot original, string expectedUid)
        {
            if (disposed || IsSyntheticDemo || !ReferenceEquals(controller.CurrentSnapshot, original) ||
                source.ActivePlanningItem.PlanningItemUID != expectedUid)
            {
                if (!disposed) owner.ShowAnalysisStatus("Plankontext wurde gewechselt; das ausstehende Analyseergebnis wurde verworfen.", MainView.AnalysisStatusSeverity.Info);
                return false;
            }
            return true;
        }

        private async void OnGenerateDrrRequested(object sender, ReviewWorkspaceActionEventArgs args)
        {
            if (IsSyntheticDemo || disposed || analysisCancellation != null || imageCancellation != null) return;
            var plan = source.ActivePlanningItem.PlanningItemObject as PlanSetup;
            var selection = CurrentViewModel.Analysis;
            if (plan == null || selection.SelectedBeam == null || !selection.SelectedBeam.BeamNumber.HasValue || selection.SelectedControlPoint == null) return;
            var original = controller.CurrentSnapshot;
            if (!CanExportCurrentSnapshot(original)) { owner.ShowAnalysisStatus("Plankontext veraltet; bitte Ansicht aktualisieren.", MainView.AnalysisStatusSeverity.Warning); return; }
            string expectedUid = plan.UID;
            int beamNumber = selection.SelectedBeam.BeamNumber.Value, cpIndex = selection.SelectedControlPoint.Index;
            analysisCancellation = new CancellationTokenSource(TimeSpan.FromSeconds(45));
            view.IsEnabled = false;
            owner.ShowAnalysisStatus("DRR aus der aktiven ESAPI-Planungs-CT wird berechnet …", MainView.AnalysisStatusSeverity.Info);
            try
            {
                var image = await new EsapiBevBuilder().BuildAsync(plan, original.PlanAnalysis, beamNumber, cpIndex, analysisCancellation.Token);
                analysisCancellation.Token.ThrowIfCancellationRequested();
                if (!AnalysisContextIsCurrent(original, expectedUid) || !CanExportCurrentSnapshot(original)) return;
                var cp = original.PlanAnalysis.Beams.Single(b => b.BeamNumber == beamNumber).ControlPoints.Single(c => c.Index == cpIndex);
                cp.BevImage = image;
                selection.RenderBevPreview();
                owner.ShowAnalysisStatus(image.SourceStatus == "available" ? "DRR aus ESAPI berechnet." : image.UnavailableReason,
                    image.SourceStatus == "available" ? MainView.AnalysisStatusSeverity.Success : MainView.AnalysisStatusSeverity.Warning,
                    "Darstellung ist eine Forschungsübersicht, keine Kollisions- oder Bildregistrierungsprüfung.");
            }
            catch (OperationCanceledException) { owner.ShowAnalysisStatus("DRR abgebrochen / Zeitlimit erreicht; keine unvollständige Projektion übernommen.", MainView.AnalysisStatusSeverity.Warning); }
            catch (Exception) { owner.ShowAnalysisStatus("DRR nicht verfügbar. Aktive Planungs-CT, Bildorientierung und Plankontext prüfen.", MainView.AnalysisStatusSeverity.Error); }
            finally { analysisCancellation.Dispose(); analysisCancellation = null; if (!disposed) view.IsEnabled = true; }
        }

        internal async Task<ClearPlan.Core.PlanAnalysis.ReviewPlanAnalysis> PrepareReportBevsAsync(ClearPlan.Core.Review.ReviewSnapshot snapshot)
        {
            var pendingImages = imageCapture;
            if (pendingImages != null && !pendingImages.IsCompleted) await pendingImages;
            if (analysisCancellation != null || !CanExportCurrentSnapshot(snapshot))
                throw new InvalidOperationException("Current plan analysis is busy or stale.");
            if (snapshot.PlanAnalysis != null && snapshot.PlanAnalysis.Beams.Count > 0 && snapshot.PlanAnalysis.Beams.All(beam =>
                beam.ControlPoints.Count(cp => cp.Index == 0 && cp.BevImage != null && cp.BevImage.SourceStatus == "available") == 1))
                return ClearPlan.Core.PlanAnalysis.PlanAnalysisSnapshot.Copy(snapshot.PlanAnalysis);
            var plan = source.ActivePlanningItem.PlanningItemObject as PlanSetup;
            if (plan == null) return snapshot.PlanAnalysis;
            string expectedUid = plan.UID;
            // The builder owns the bounded capture budget and marks timed-out optional images
            // unavailable. This outer source is only for explicit/context cancellation.
            analysisCancellation = new CancellationTokenSource();
            view.IsEnabled = false;
            try
            {
                await System.Windows.Threading.Dispatcher.Yield(System.Windows.Threading.DispatcherPriority.Background);
                if (!CanExportCurrentSnapshot(snapshot)) throw new InvalidOperationException("Plan context changed before report capture.");
                owner.ShowAnalysisStatus("Report: DRRs der Feldstarts aus der aktuellen ESAPI-CT werden berechnet …", MainView.AnalysisStatusSeverity.Info);
                var result = await new EsapiBevBuilder().BuildStartsAsync(plan, snapshot.PlanAnalysis, analysisCancellation.Token);
                analysisCancellation.Token.ThrowIfCancellationRequested();
                if (!AnalysisContextIsCurrent(snapshot, expectedUid) || !CanExportCurrentSnapshot(snapshot))
                    throw new InvalidOperationException("Plan changed during report image capture.");
                foreach (var beam in result.Beams)
                {
                    var targetBeam = snapshot.PlanAnalysis.Beams.SingleOrDefault(row => row.BeamNumber == beam.BeamNumber);
                    var first = beam.ControlPoints.SingleOrDefault(cp => cp.Index == 0);
                    var target = targetBeam == null ? null : targetBeam.ControlPoints.SingleOrDefault(cp => cp.Index == 0);
                    if (target != null && first != null) target.BevImage = first.BevImage;
                }
                CurrentViewModel.Analysis.RenderBevPreview();
                return result;
            }
            finally { analysisCancellation.Dispose(); analysisCancellation = null; if (!disposed) view.IsEnabled = true; }
        }

        public bool CanExportCurrentSnapshot(ClearPlan.Core.Review.ReviewSnapshot snapshot)
        {
            return !disposed && !IsSyntheticDemo && ReferenceEquals(controller.CurrentSnapshot, snapshot) &&
                source.ActivePlanningItem != null && planContext.Matches(snapshot, source.ActivePlanningItem.PlanningItemUID, source.ActivePlanningItem.PlanningItemObject);
        }

        public void RequestCurrentReport()
        {
            ThrowIfDisposed();
            if (CurrentViewModel == null)
            {
                owner.ShowAnalysisStatus("Report nicht verfügbar: aktuelle Gesamtansicht konnte nicht geladen werden.", MainView.AnalysisStatusSeverity.Error);
                return;
            }
            OnReportRequested(CurrentViewModel, null);
        }

        private void OnReportRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (IsSyntheticDemo)
            {
                owner.HandleSharedSyntheticReport(
                    controller.CurrentSnapshot, (ReviewWorkspaceViewModel)sender);
                return;
            }

            if (!CanExportCurrentSnapshot(controller.CurrentSnapshot))
            {
                owner.ShowAnalysisStatus("Report abgebrochen: Plankontext stimmt nicht mehr mit der geprüften Ansicht überein.", MainView.AnalysisStatusSeverity.Error);
                return;
            }
            owner.HandleSharedPlanReport(controller.CurrentSnapshot,
                (ReviewWorkspaceViewModel)sender);
        }

        private void OnHtmlReportRequested(object sender, ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (!ReferenceEquals(sender, CurrentViewModel)) return;
            HtmlReportTask = PrepareHtmlQuicklookAsync();
        }

        private void OnWorkspaceDataContextChanged(object sender,System.Windows.DependencyPropertyChangedEventArgs args)
        {
            if (collisionCancellation != null) collisionCancellation.Cancel();
            owner.CancelAriaPreparation();
            owner.RefreshAriaAvailability(controller.CurrentViewModel);
        }

        private void OnAriaUploadRequested(object sender,ReviewWorkspaceActionEventArgs args)
        {
            if(disposed || IsSyntheticDemo || !ReferenceEquals(sender,CurrentViewModel) || !CanExportCurrentSnapshot(controller.CurrentSnapshot)) return;
            AriaUploadTask=owner.HandleSharedAriaUploadAsync(controller.CurrentSnapshot,CurrentViewModel);
        }

        private async Task PrepareHtmlQuicklookAsync()
        {
            HtmlReportDiagnostic = null;
            var snapshot = controller.CurrentSnapshot;
            var workspace = CurrentViewModel;
            try
            {
                // Quicklook reuses the snapshot; only explicitly enabled, missing
                // field-start DRRs are acquired, never goals/DVH/PAM recalculated.
                if (!IsSyntheticDemo && workspace.IncludeBeamEyeViews)
                    await PrepareReportBevsAsync(snapshot);
                if (disposed || !ReferenceEquals(workspace, CurrentViewModel)) return;
                owner.HandleSharedHtmlReport(snapshot, workspace);
                HtmlReportDiagnostic = owner.HtmlReportDiagnostic;
            }
            catch (Exception exception)
            {
                HtmlReportDiagnostic = ReportDiagnostic(exception);
                if (!disposed) owner.ShowAnalysisStatus("HTML-Quicklook nicht erstellt: Plananalyse läuft oder Bildaufnahme nicht möglich. Erneut versuchen oder BEV im Report abwählen.", MainView.AnalysisStatusSeverity.Warning);
            }
        }

        internal string HtmlReportDiagnostic { get; private set; }
        internal static string ReportDiagnostic(Exception exception)
        {
            // No exception messages, filenames, arguments or patient data.
            return exception.GetType().FullName + " | " + string.Join(" <- ",
                (new System.Diagnostics.StackTrace(exception).GetFrames() ?? new System.Diagnostics.StackFrame[0])
                .Select(frame => frame.GetMethod()).Where(method => method != null)
                .Select(method => (method.DeclaringType == null ? "" : method.DeclaringType.FullName) + "." + method.Name).Take(10));
        }

        private void OnOpenPlanRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (IsSyntheticDemo)
            {
                owner.ShowSyntheticDemoActionNotice(
                    "Synthetische Pläne werden nicht in Eclipse geöffnet.");
                return;
            }

            owner.HandleSharedOpenPlan(
                eventArgs.Parameter as ReviewPlanRowViewModel);
        }

        private void OnComparisonPlanRequested(object sender, ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (!ReferenceEquals(sender, CurrentViewModel) || disposed || analysisCancellation != null || imageCancellation != null) return;
            ComparisonTask = LoadComparisonPlanAsync(eventArgs.Parameter as ReviewPlanRowViewModel);
        }

        private async Task LoadComparisonPlanAsync(ReviewPlanRowViewModel row)
        {
            var workspace = CurrentViewModel;
            var original = controller.CurrentSnapshot;
            if (row == null || IsSyntheticDemo || !CanExportCurrentSnapshot(original)) return;
            var selected = owner.FindSharedPlanningItem(row);
            if (selected == null) { workspace.Comparison.SetLoading(false, "Referenzplan nicht eindeutig verfügbar. Planauswahl aktualisieren."); return; }
            if (row.PlanKey == original.ActivePlanKey) { workspace.Comparison.SetReference(original, true); return; }
            analysisCancellation = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            view.IsEnabled = false;
            workspace.Comparison.SetLoading(true, "Referenzplan wird lesend geladen; der aktuelle Plan bleibt geöffnet …");
            try
            {
                await System.Windows.Threading.Dispatcher.Yield(System.Windows.Threading.DispatcherPriority.Background);
                if (!CanExportCurrentSnapshot(original)) return;
                // A separate adapter prevents reference-table evaluation from replacing
                // the active plan's constraints, mappings, dose presentation or GUI.
                var referenceSource = new MainViewModel(source.User, source.Patient, source.ScriptVersion,
                    source.PlanningItemList, selected, selected.PlanningItemObject);
                referenceSource.GetPQMSummaries(referenceSource.ActiveConstraintPath, selected, source.Patient);
                var snapshot = new EsapiReviewSnapshotBuilder().Build(referenceSource, settingsProvider());
                var plan = selected.PlanningItemObject as PlanSetup;
                if (plan != null)
                {
                    var builder = new EsapiPlanAnalysisBuilder(nativeProfiles, doseRateProfiles);
                    snapshot.PlanAnalysis = await builder.EnrichTargetProjectionsAsync(plan, builder.Build(plan), analysisCancellation.Token);
                }
                analysisCancellation.Token.ThrowIfCancellationRequested();
                if (!CanExportCurrentSnapshot(original) || !ReferenceEquals(workspace, CurrentViewModel)) return;
                if (!ClearPlan.Core.Review.ReviewSnapshotValidator.Validate(snapshot).IsValid)
                    throw new InvalidOperationException("Reference snapshot validation failed.");
                workspace.Comparison.SetReference(snapshot, true);
            }
            catch (OperationCanceledException) { workspace.Comparison.SetLoading(false, "Referenz nicht übernommen: Zeitlimit oder Abbruch. Der bisherige Vergleich bleibt erhalten."); }
            catch (Exception) { workspace.Comparison.SetLoading(false, "Referenz konnte nicht geladen werden. Plankontext und Datenverfügbarkeit prüfen; bisheriger Vergleich bleibt erhalten."); }
            finally { workspace.Comparison.SetLoading(false, null); analysisCancellation.Dispose(); analysisCancellation = null; if (!disposed) view.IsEnabled = true; }
        }

        private void OnDvhResetRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (IsSyntheticDemo)
            {
                ((ReviewWorkspaceViewModel)sender).ResetDvhSelections();
                return;
            }

            owner.HandleSharedDvhReset(
                (ReviewWorkspaceViewModel)sender);
        }

        private void OnDvhExportRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (IsSyntheticDemo)
            {
                owner.ShowSyntheticDemoActionNotice(
                    "Für reproduzierbare DVH-Dateien bitte den Standalone-Simulator verwenden.");
                return;
            }

            owner.HandleSharedDvhExport(
                (ReviewWorkspaceViewModel)sender);
        }

        private void OnNavigationRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (string.Equals(eventArgs.Parameter as string, "PlanImagesTab", StringComparison.Ordinal))
                QueuePlanImages(false);
            if (string.Equals(eventArgs.Parameter as string, "CollisionTab", StringComparison.Ordinal))
                QueueCollision(false);
            owner.HandleSharedNavigation(
                eventArgs.Parameter as string);
        }

        private void OnPlanImagesRequested(object sender, ReviewWorkspaceActionEventArgs args)
        {
            QueuePlanImages(true);
        }

        private void QueuePlanImages(bool force)
        {
            if (disposed || CurrentViewModel == null) return;
            if (IsSyntheticDemo)
            {
                CurrentViewModel.PlanImages.ReplaceImages(SyntheticPlanImageFactory.Create(CurrentViewModel.ActivePlanKey));
                return;
            }
            if (!force && CurrentViewModel.PlanImages.HasImages) return;
            var previous = PlanImagesTask ?? Task.FromResult(0);
            PlanImagesTask = LoadCurrentPlanImagesAsync(force, source.ActivePlanningItem, previous);
        }

        private void QueueImagesIfVisible()
        {
            var tab = view.FindName("PlanImagesTab") as System.Windows.Controls.TabItem;
            if (tab != null && tab.IsSelected) QueuePlanImages(false);
        }

        private async Task LoadCurrentPlanImagesAsync(bool force, PlanningItemViewModel expectedPlanningItem, Task previous)
        {
            CurrentViewModel.PlanImages.SetLoading("Schnittebenen werden vorbereitet …");
            try
            {
                try { await previous; } catch (OperationCanceledException) { }
                // Parameter analysis replaces the initial detached snapshot. Finish it first,
                // then capture against the newly committed context, never a stale snapshot.
                var analysis = NativeAnalysisTask;
                if (analysis != null) await analysis;
                var previousCapture = imageCapture;
                if (previousCapture != null && !previousCapture.IsCompleted)
                {
                    // A previous failed/canceled read must not poison an explicit retry.
                    // The fresh capture below still reports its own failure to this view.
                    try { await previousCapture; } catch (Exception) { }
                }
                if (disposed || IsSyntheticDemo || !ReferenceEquals(source.ActivePlanningItem, expectedPlanningItem)) return;
                var snapshot = controller.CurrentSnapshot;
                if (force) snapshot.PlanImages = new System.Collections.Generic.List<ClearPlan.Core.Review.ReviewPlanImage>();
                await PreparePlanImagesAsync(snapshot);
            }
            catch (OperationCanceledException) { }
            catch (Exception)
            {
                if (!disposed && !IsSyntheticDemo && ReferenceEquals(source.ActivePlanningItem, expectedPlanningItem))
                    CurrentViewModel.PlanImages.SetFailure("Schnittbilder konnten nicht geladen werden. Erneut laden oder aktuellen Plan prüfen.");
            }
        }

        internal Task<System.Collections.Generic.List<ClearPlan.Core.Review.ReviewPlanImage>> PreparePlanImagesAsync(ClearPlan.Core.Review.ReviewSnapshot snapshot)
        {
            if (!CanExportCurrentSnapshot(snapshot)) throw new InvalidOperationException("CT capture requires the current native plan context.");
            if (snapshot.PlanImages != null && snapshot.PlanImages.Count == 3 && snapshot.PlanImages.All(i => i.PlanKey == snapshot.ActivePlanKey && !i.Synthetic))
            {
                CurrentViewModel.PlanImages.ReplaceImages(snapshot.PlanImages);
                return Task.FromResult(snapshot.PlanImages);
            }
            if (imageCapture != null && !imageCapture.IsCompleted)
            {
                if (ReferenceEquals(imageCaptureSnapshot, snapshot)) return imageCapture;
                throw new InvalidOperationException("Previous CT capture is still completing.");
            }
            if (analysisCancellation != null) throw new InvalidOperationException("Native analysis is busy; CT capture will be available afterwards.");
            imageCaptureSnapshot = snapshot;
            imageCapture = CapturePlanImagesAsync(snapshot);
            return imageCapture;
        }

        private async Task<System.Collections.Generic.List<ClearPlan.Core.Review.ReviewPlanImage>> CapturePlanImagesAsync(ClearPlan.Core.Review.ReviewSnapshot snapshot)
        {
            var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(65));
            imageCancellation = cancellation;
            CurrentViewModel.PlanImages.SetLoading("CT, Strukturkonturen und Isodosen werden gelesen …");
            try
            {
                // Explicit runner actions can start on the STA without a WPF synchronization
                // context. Enter its dispatcher before the first async native capture; never
                // fall back to a worker or carry a native object across this initial yield.
                await System.Windows.Threading.Dispatcher.Yield(System.Windows.Threading.DispatcherPriority.Background);
                cancellation.Token.ThrowIfCancellationRequested();
                if (!CanExportCurrentSnapshot(snapshot)) throw new OperationCanceledException();
                var images = await EsapiPlanImageBuilder.BuildAsync(source.ActivePlanningItem.PlanningItemObject,
                    snapshot.ActivePlanKey, cancellation.Token, () => CanExportCurrentSnapshot(snapshot));
                cancellation.Token.ThrowIfCancellationRequested();
                if (!CanExportCurrentSnapshot(snapshot)) throw new OperationCanceledException();
                var previous = snapshot.PlanImages;
                snapshot.PlanImages = images;
                if (!ClearPlan.Core.Review.ReviewSnapshotValidator.Validate(snapshot).IsValid)
                {
                    snapshot.PlanImages = previous;
                    throw new InvalidOperationException("CT capture did not satisfy snapshot validation.");
                }
                CurrentViewModel.PlanImages.ReplaceImages(images);
                return images;
            }
            catch (OperationCanceledException)
            {
                if (CanExportCurrentSnapshot(snapshot)) CurrentViewModel.PlanImages.SetFailure("Bildaufnahme abgebrochen. Erneut laden.");
                throw;
            }
            catch (Exception)
            {
                if (CanExportCurrentSnapshot(snapshot)) CurrentViewModel.PlanImages.SetFailure("CT oder Overlays nicht verfügbar. Erneut laden oder Datenquelle prüfen.");
                throw;
            }
            finally
            {
                if (ReferenceEquals(imageCancellation, cancellation)) imageCancellation = null;
                cancellation.Dispose();
            }
        }

        private void OnStructureMappingRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (IsSyntheticDemo)
            {
                var workspace = (ReviewWorkspaceViewModel)sender;
                var mapping = eventArgs.Parameter as
                    ReviewStructureMappingViewModel;
                if (mapping != null &&
                    !string.IsNullOrWhiteSpace(mapping.SelectedStructureId))
                {
                    workspace.MarkStructureMappingApplied(
                        mapping.StableId,
                        mapping.SelectedStructureId,
                        "Synthetic mapping applied for demonstration only.");
                }
                return;
            }

            owner.HandleSharedStructureMapping(
                (ReviewWorkspaceViewModel)sender,
                eventArgs.Parameter as ReviewStructureMappingViewModel);
        }

        private void OnConstraintTableRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (IsSyntheticDemo)
            {
                owner.ShowSyntheticDemoActionNotice(
                    "Im Demomodus ist die synthetische Prüftabelle fest vorgegeben.");
                return;
            }

            owner.HandleSharedConstraintSelection();
        }

        private void OnSettingsRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            owner.HandleSharedSettings();
        }

        private void OnScenarioSelectionRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            // Clinical mode has no scenario switch. The event remains
            // subscribed so the shared controller has one lifecycle contract.
        }

        private void ThrowIfDisposed()
        {
            if (disposed)
            {
                throw new ObjectDisposedException(
                    typeof(ClinicalReviewWorkspaceHost).FullName);
            }
        }
    }
}
