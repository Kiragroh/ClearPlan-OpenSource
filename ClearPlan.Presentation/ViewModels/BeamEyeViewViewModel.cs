using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Simulation;
using ClearPlan.Rendering;

namespace ClearPlan.Presentation.ViewModels
{
    public sealed partial class PlanAnalysisViewModel
    {
        private int controlPointPosition;
        private bool bevActive;
        private bool bevSynthetic;
        private bool generatingDrr;
        private bool initialNativeDrrRequested;
        private Task syntheticDrrTask;
        private readonly Func<ReviewBeamAnalysis, ReviewControlPointSample, BeamEyeViewImage> syntheticDrrProvider;
        public ImageSource BevImageSource { get; private set; }
        public ICommand GenerateDrrCommand { get; internal set; }
        public ICommand PreviousControlPointCommand { get; private set; }
        public ICommand NextControlPointCommand { get; private set; }
        public string BevStatusText { get; private set; }
        public string BevStatusSummary { get; private set; }
        public int MaximumControlPointPosition { get { return Math.Max(0, selectedBeam == null || selectedBeam.ControlPoints == null ? 0 : selectedBeam.ControlPoints.Count - 1); } }
        public ReviewControlPointSample SelectedControlPoint
        {
            get { return selectedBeam == null || selectedBeam.ControlPoints == null || selectedBeam.ControlPoints.Count == 0 ? null : selectedBeam.ControlPoints[controlPointPosition]; }
        }
        public int ControlPointPosition
        {
            get { return controlPointPosition; }
            set
            {
                controlPointPosition = Math.Max(0, Math.Min(MaximumControlPointPosition, value));
                Changed("ControlPointPosition"); Changed("SelectedControlPoint"); Changed("ControlPointText");
                if (bevActive) RenderBevPreview();
                CommandManager.InvalidateRequerySuggested();
            }
        }
        public string ControlPointText { get { return SelectedControlPoint == null ? "Keine Kontrollpunkte" : "CP " + SelectedControlPoint.Index + " · " + (controlPointPosition + 1) + " / " + (MaximumControlPointPosition + 1) + (controlPointPosition == 0 ? " · Feldstart" : ""); } }

        private void InitializeBev(bool synthetic)
        {
            bevSynthetic = synthetic;
            PreviousControlPointCommand = new RelayCommand(p => ControlPointPosition--, p => ControlPointPosition > 0);
            NextControlPointCommand = new RelayCommand(p => ControlPointPosition++, p => ControlPointPosition < MaximumControlPointPosition);
            GenerateDrrCommand = new RelayCommand(async p => await GenerateSyntheticDrrAsync(), p => bevSynthetic && SelectedControlPoint != null && !generatingDrr);
        }

        private void ResetBevSelection()
        {
            controlPointPosition = 0;
            Changed("MaximumControlPointPosition");
            ControlPointPosition = 0;
        }

        public async Task ActivateBevAsync()
        {
            bevActive = true;
            RenderBevPreview();
            if (bevSynthetic && SelectedControlPoint != null && SelectedControlPoint.BevImage == null)
                await GenerateSyntheticDrrAsync();
            else if (!bevSynthetic && !initialNativeDrrRequested && SelectedControlPoint != null &&
                (SelectedControlPoint.BevImage == null ||
                 string.Equals(SelectedControlPoint.BevImage.SourceStatus, "available", StringComparison.OrdinalIgnoreCase) &&
                 !BeamEyeViewRenderer.InspectState(selectedBeam, SelectedControlPoint, SelectedControlPoint.BevImage, false).ImageAvailable) &&
                GenerateDrrCommand != null && GenerateDrrCommand.CanExecute(null))
            {
                // The host owns ESAPI capture, cancellation and snapshot freshness. One initial
                // request per view model; CP navigation and explicit retries remain deliberate.
                initialNativeDrrRequested = true;
                GenerateDrrCommand.Execute(null);
            }
        }

        public void RenderBevPreview()
        {
            var cp = SelectedControlPoint;
            byte[] bytes = BeamEyeViewRenderer.Render(selectedBeam, cp, cp == null ? null : cp.BevImage, bevSynthetic, compactStatus: true);
            using (var stream = new MemoryStream(bytes, false))
            {
                var image = new BitmapImage();
                image.BeginInit(); image.CacheOption = BitmapCacheOption.OnLoad; image.StreamSource = stream; image.EndInit(); image.Freeze();
                BevImageSource = image;
            }
            var state = BeamEyeViewRenderer.InspectState(selectedBeam, cp, cp == null ? null : cp.BevImage, bevSynthetic);
            BevStatusSummary = state.ImageAvailable ? "CT-Projektion verfügbar · Details" : "CT-Projektion fehlt · Details";
            BevStatusText = state.ImageAvailable ? cp.BevImage.ProjectionDescription :
                cp != null && cp.BevImage == null ?
                "Für diesen Kontrollpunkt ist noch keine CT-Projektion geladen. Es wird kein anderes Projektionsbild unterlegt." :
                state.ImageMessage;
            if (!state.GeometryAvailable) BevStatusText += "\nMLC / Blenden: " + state.GeometryMessage;
            if (!state.ImageAvailable && cp != null) BevStatusText += "\nMit „DRR laden“ die CT-Projektion für diesen Kontrollpunkt anfordern.";
            Changed("BevImageSource"); Changed("BevStatusText"); Changed("BevStatusSummary");
        }

        private Task GenerateSyntheticDrrAsync()
        {
            if (syntheticDrrTask != null && !syntheticDrrTask.IsCompleted) return syntheticDrrTask;
            syntheticDrrTask = GenerateSyntheticDrrCoreAsync();
            return syntheticDrrTask;
        }

        private async Task GenerateSyntheticDrrCoreAsync()
        {
            if (!bevSynthetic || generatingDrr || SelectedControlPoint == null) return;
            var beam = selectedBeam; var cp = SelectedControlPoint;
            generatingDrr = true;
            BevStatusSummary = "CT-Projektion wird berechnet · Details";
            BevStatusText = "Synthetische CT-Projektion wird berechnet …"; Changed("BevStatusText"); Changed("BevStatusSummary");
            try
            {
                // Detached synthetic DTOs only; this path has no vendor or patient access.
                var image = await Task.Run(() => syntheticDrrProvider(beam, cp));
                cp.BevImage = image;
                if (ReferenceEquals(selectedBeam, beam) && ReferenceEquals(SelectedControlPoint, cp)) RenderBevPreview();
            }
            catch (Exception)
            {
                BevStatusSummary = "CT-Projektion konnte nicht berechnet werden · Details";
                BevStatusText = "Synthetische DRR konnte nicht berechnet werden. Mit „DRR laden“ erneut versuchen.";
                Changed("BevStatusText"); Changed("BevStatusSummary");
            }
            finally { generatingDrr = false; CommandManager.InvalidateRequerySuggested(); }
        }
    }
}
