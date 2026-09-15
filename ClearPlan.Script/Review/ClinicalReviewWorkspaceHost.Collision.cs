using System;
using System.Threading;
using System.Threading.Tasks;
using ClearPlan.Core.Collision;
using ClearPlan.Presentation.ViewModels;
using VMS.TPS.Common.Model.API;

namespace ClearPlan.Review
{
    internal sealed partial class ClinicalReviewWorkspaceHost
    {
        private CancellationTokenSource collisionCancellation;
        internal Task CollisionTask { get; private set; }
        private void OnCollisionRequested(object sender, ReviewWorkspaceActionEventArgs args) { QueueCollision(true); }
        private void QueueCollision(bool force)
        {
            if (disposed || CurrentViewModel == null || collisionCancellation != null) return;
            if (!force && CurrentViewModel.Collision.HasScene) return;
            if (IsSyntheticDemo) { CurrentViewModel.Collision.ReloadCommand.Execute(null); return; }
            if (analysisCancellation != null || imageCancellation != null)
            { CurrentViewModel.Collision.SetFailure("Plananalyse läuft noch. Anschließend Geometrie laden."); return; }
            CollisionTask = CaptureCollisionAsync();
        }
        private async Task CaptureCollisionAsync()
        {
            var snapshot = controller.CurrentSnapshot;
            var workspace = CurrentViewModel;
            bool zeroPitchRollConfirmed = workspace.Collision.ZeroPitchRollConfirmed;
            var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(40));
            collisionCancellation = cancellation;
            workspace.Collision.SetLoading();
            try
            {
                await System.Windows.Threading.Dispatcher.Yield(System.Windows.Threading.DispatcherPriority.Background);
                if (!CanExportCurrentSnapshot(snapshot)) throw new OperationCanceledException();
                var settings = settingsProvider();
                string path = settings.Paths.CollisionProfilesJsonPath;
                CollisionProfileCatalog profiles = null;
                try { profiles = await CollisionProfileFileSource.LoadAsync(string.IsNullOrWhiteSpace(path) ? null : settings.ResolvePath(path), cancellation.Token); }
                catch (OperationCanceledException) { throw; }
                catch (Exception) { /* Surfaces remain inspectable; unavailable profile never becomes clearance. */ }
                SourceCollisionModelCatalog sourceModels = null;
                string sourcePath = settings.Paths.SourceCollisionModelsJsonPath;
                try { sourceModels = await CollisionProfileFileSource.LoadSourceModelsAsync(string.IsNullOrWhiteSpace(sourcePath) ? null : settings.ResolvePath(sourcePath), cancellation.Token); }
                catch (OperationCanceledException) { throw; }
                catch (Exception) { /* A source model is optional and is never a commissioned-envelope fallback. */ }
                cancellation.Token.ThrowIfCancellationRequested();
                if (!CanExportCurrentSnapshot(snapshot) || !ReferenceEquals(CurrentViewModel, workspace)) throw new OperationCanceledException();
                // Vendor and WPF mesh access remains synchronous on this owner dispatcher.
                var scene = EsapiCollisionBuilder.Capture(source.ActivePlanningItem.PlanningItemObject as PlanSetup,
                    snapshot.ActivePlanKey, profiles, zeroPitchRollConfirmed, cancellation.Token);
                var evaluation = await Task.Run(() => {
                    SourceCollisionSceneAdapter.Attach(scene, sourceModels, cancellation.Token);
                    return new { Result = CollisionEvaluator.Evaluate(scene, cancellation.Token),
                        Sweep = CollisionSweep.Evaluate(scene, cancellation.Token) };
                }, cancellation.Token);
                cancellation.Token.ThrowIfCancellationRequested();
                if (!CanExportCurrentSnapshot(snapshot) || !ReferenceEquals(CurrentViewModel, workspace) ||
                    workspace.Collision.ZeroPitchRollConfirmed != zeroPitchRollConfirmed) throw new OperationCanceledException();
                workspace.Collision.SetScene(scene, evaluation.Result, evaluation.Sweep);
            }
            catch (OperationCanceledException)
            { if (!disposed && ReferenceEquals(CurrentViewModel, workspace)) workspace.Collision.SetFailure("Geometrieprüfung abgebrochen oder Zeitlimit erreicht. Erneut laden."); }
            catch (Exception)
            { if (!disposed && ReferenceEquals(CurrentViewModel, workspace)) workspace.Collision.SetFailure("Geometrie nicht bewertbar. Oberflächen, Geräteprofil und unterstützte Lagerung prüfen."); }
            finally { if (ReferenceEquals(collisionCancellation, cancellation)) collisionCancellation = null; cancellation.Dispose(); }
        }
    }
}
