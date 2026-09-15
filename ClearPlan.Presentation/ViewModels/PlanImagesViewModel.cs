using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ClearPlan.Core.Review;
using ClearPlan.Rendering;

namespace ClearPlan.Presentation.ViewModels
{
    /// <summary>Read-only orthogonal overview. No volume navigation or native planning objects are retained.</summary>
    public sealed partial class PlanImagesViewModel : INotifyPropertyChanged
    {
        private readonly string planKey;
        private List<ReviewPlanImage> sourceImages = new List<ReviewPlanImage>();
        private bool showStructures = true, showDose = true, focusIsocenter = true;
        private string enlargedKind;
        private readonly Dictionary<string, bool> structureVisibility = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        public event PropertyChangedEventHandler PropertyChanged;
        public event EventHandler ReloadRequested;
        public event EventHandler<StructureVisibilityEventArgs> StructureVisibilityChanged;
        public bool IsStructureVisible(string id) { bool value; return !structureVisibility.TryGetValue(id ?? "", out value) || value; }
        public void SetStructureVisibility(string id, bool visible)
        {
            if (string.IsNullOrWhiteSpace(id)) return;
            bool old;
            if (structureVisibility.TryGetValue(id, out old) && old == visible) return;
            structureVisibility[id] = visible;
            if (sourceImages.Count != 0) RebuildPlanes();
        }
        public void SetStructureVisibilities(IEnumerable<KeyValuePair<string, bool>> values)
        {
            foreach (var pair in values) structureVisibility[pair.Key] = pair.Value;
            if (sourceImages.Count != 0) RebuildPlanes();
        }

        public PlanImagesViewModel(IEnumerable<ReviewPlanImage> images, string planKey)
        {
            this.planKey = planKey;
            InitializeIsodoses();
            EnlargePlaneCommand = new RelayCommand(p =>
            {
                string kind = p as string;
                if (Planes.Any(plane => plane.Kind == kind)) { enlargedKind = kind; SelectionChanged(); }
            });
            ShowAllPlanesCommand = new RelayCommand(p => { enlargedKind = null; SelectionChanged(); }, p => IsEnlarged);
            FitCommand = new RelayCommand(p => FocusIsocenter = false, p => HasImages);
            FocusIsocenterCommand = new RelayCommand(p => FocusIsocenter = true, p => HasImages);
            ReloadCommand = new RelayCommand(p =>
            {
                var handler = ReloadRequested; if (handler != null) handler(this, EventArgs.Empty);
            }, p => !IsLoading);
            ReplaceImages(images);
        }

        public IList<PlanImagePlaneViewModel> Planes { get; private set; }
        public IEnumerable<PlanImagePlaneViewModel> VisiblePlanes { get { return IsEnlarged ? Planes.Where(p => p.Kind == enlargedKind) : Planes; } }
        public IList<PlanImageLegendItem> LegendItems { get; private set; }
        public bool IsEnlarged { get { return enlargedKind != null; } }
        public string HeadingText { get { return IsEnlarged ? "Schnittbilder · " + Planes.First(p => p.Kind == enlargedKind).Title : "Schnittbilder"; } }
        public int PlaneColumns { get { return IsEnlarged ? 1 : 3; } }
        public bool HasImages { get { return Planes != null && Planes.Any(p => p.IsAvailable); } }
        public bool IsLoading { get; private set; }
        public string StatusText { get; private set; }
        public string AvailabilityText { get { return Planes.Count(p => p.IsAvailable) + " / 3 Ebenen verfügbar"; } }
        public string ViewportModeText { get { return FocusIsocenter ? "Iso-Fokus · 2%-Rx-Bereich, falls verfügbar" : "Gesamte CT-Ausdehnung"; } }
        public string LegendSummary { get { return LegendItems.Count == 0 ? "Keine eingeblendeten Konturen verfügbar" : LegendItems.Count + " Konturen / Dosislinien · Legende und Quellen"; } }
        public string WindowLevelDescription { get { return "Festes CT-Fenster: Breite 400 HU / Lage 40 HU. Die Vorschau enthält bereits gefensterte 8-Bit-Pixel; Fenster und Lage sind hier nicht nachträglich einstellbar. Bei Simulation: HU-äquivalentes mathematisches Phantom."; } }
        public bool ShowStructures { get { return showStructures; } set { if (showStructures == value) return; showStructures = value; RebuildPlanes(); Changed("ShowStructures"); } }
        public bool ShowDose { get { return showDose; } set { if (showDose == value) return; showDose = value; RebuildPlanes(); Changed("ShowDose"); } }
        public bool FocusIsocenter { get { return focusIsocenter; } set { if (focusIsocenter == value) return; focusIsocenter = value; RebuildPlanes(); Changed("FocusIsocenter"); Changed("ViewportModeText"); } }
        public ICommand EnlargePlaneCommand { get; private set; }
        public ICommand ShowAllPlanesCommand { get; private set; }
        public ICommand FitCommand { get; private set; }
        public ICommand FocusIsocenterCommand { get; private set; }
        public ICommand ReloadCommand { get; private set; }

        public void ReplaceImages(IEnumerable<ReviewPlanImage> images)
        {
            // Plan identity is strict: unrelated or unidentified captures never become a fallback image.
            sourceImages = (images ?? Enumerable.Empty<ReviewPlanImage>()).Where(image => image != null &&
                !string.IsNullOrWhiteSpace(planKey) && string.Equals(image.PlanKey, planKey, StringComparison.Ordinal)).ToList();
            IsLoading = false;
            doseDisplayDirty = true;
            RebuildPlanes();
            StatusText = HasImages ? "Orthogonale Übersichten der erfassten Ebene · nur Anzeige, keine Schichtnavigation" :
                "Für diesen Plan sind keine gültigen Schnittbilder geladen. Mit „CT laden“ erneut anfordern.";
            Changed("StatusText"); Changed("IsLoading"); CommandManager.InvalidateRequerySuggested();
        }

        public void SetLoading(string message)
        {
            sourceImages.Clear(); doseDisplayDirty = true; IsLoading = true; RebuildPlanes();
            StatusText = string.IsNullOrWhiteSpace(message) ? "Schnittbilder werden geladen …" : message;
            Changed("StatusText"); Changed("IsLoading"); CommandManager.InvalidateRequerySuggested();
        }

        public void SetFailure(string message)
        {
            sourceImages.Clear(); doseDisplayDirty = true; IsLoading = false; RebuildPlanes();
            StatusText = string.IsNullOrWhiteSpace(message) ? "Schnittbilder konnten nicht geladen werden. Mit „CT laden“ erneut versuchen." : message;
            Changed("StatusText"); Changed("IsLoading"); CommandManager.InvalidateRequerySuggested();
        }

        private void RebuildPlanes()
        {
            RefreshDoseDisplay();
            string[] kinds = { "transversal", "coronal", "sagittal" };
            string[] titles = { "Transversal", "Koronal", "Sagittal" };
            Planes = kinds.Select((kind, index) =>
            {
                var candidates = doseDisplayImages.Where(image => image.Kind == kind).ToList();
                return new PlanImagePlaneViewModel(kind, titles[index], candidates.Count == 1 ? candidates[0] : null,
                    candidates.Count > 1 ? "Mehrere Bildstände dieser Ebene; keine eindeutige Zuordnung." :
                    IsLoading ? "Bild wird geladen …" : "Kein Bild für diese Ebene geladen.", ShowStructures, ShowDose, FocusIsocenter,
                    structureVisibility.Where(pair => !pair.Value).Select(pair => pair.Key));
            }).ToList().AsReadOnly();
            LegendItems = Planes.Where(plane => plane.IsAvailable).SelectMany(plane => plane.LegendItems)
                .GroupBy(item => item.Kind + "\n" + item.Label + "\n" + item.ColorHex)
                .Select(group => group.First()).ToList().AsReadOnly();
            foreach (var item in LegendItems)
            {
                item.IsVisible = item.Kind != "structure" || IsStructureVisible(item.Label);
                item.PropertyChanged += (sender, args) =>
                {
                    var selected = (PlanImageLegendItem)sender;
                    if (args.PropertyName != "IsVisible" || selected.Kind != "structure") return;
                    SetStructureVisibility(selected.Label, selected.IsVisible);
                    var handler = StructureVisibilityChanged;
                    if (handler != null) handler(this, new StructureVisibilityEventArgs(selected.Label, selected.IsVisible));
                };
            }
            Changed("Planes"); Changed("VisiblePlanes"); Changed("HasImages"); Changed("AvailabilityText");
            Changed("LegendItems"); Changed("LegendSummary"); CommandManager.InvalidateRequerySuggested();
            Changed("IsodoseScaleItems"); Changed("IsodoseAvailabilityText");
            Changed("HasIsodoseScale"); Changed("StructureLegendItems");
        }

        private void SelectionChanged()
        { Changed("IsEnlarged"); Changed("HeadingText"); Changed("PlaneColumns"); Changed("VisiblePlanes"); CommandManager.InvalidateRequerySuggested(); }
        private void Changed(string name)
        { var handler = PropertyChanged; if (handler != null) handler(this, new PropertyChangedEventArgs(name)); }
    }

    public sealed class PlanImagePlaneViewModel
    {
        internal PlanImagePlaneViewModel(string kind, string title, ReviewPlanImage image, string missingReason,
            bool showStructures, bool showDose, bool focusIsocenter, IEnumerable<string> hiddenIds)
        {
            Kind = kind; Title = title; LegendItems = new List<PlanImageLegendItem>().AsReadOnly();
            AvailabilityText = missingReason;
            if (image == null) return;
            if (!PlanImageRenderer.IsRenderable(image))
            {
                AvailabilityText = string.IsNullOrWhiteSpace(image.UnavailableReason) ? "CT nicht verfügbar oder Bildgeometrie unvollständig." : image.UnavailableReason;
                return;
            }
            try
            {
                var hidden = new HashSet<string>(hiddenIds, StringComparer.OrdinalIgnoreCase);
                using (var stream = new MemoryStream(PlanImageRenderer.RenderFiltered(image, showStructures, showDose, focusIsocenter, hidden, false), false))
                {
                    var bitmap = new BitmapImage(); bitmap.BeginInit(); bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.StreamSource = stream; bitmap.EndInit(); bitmap.Freeze(); ImageSource = bitmap;
                }
                IsAvailable = true;
                var overlays = (image.Overlays ?? new List<ReviewImageOverlay>()).Where(overlay => overlay != null).ToList();
                var displayed = overlays.Where(overlay => overlay.SourceStatus == ReviewStatusCodes.Available &&
                    ((showDose && overlay.Kind == "isodose") || (showStructures && overlay.Kind == "structure")) &&
                    overlay.Paths != null && overlay.Paths.Any(path => path != null && path.Points != null && path.Points.Count >= 2)).ToList();
                LegendItems = displayed.Select(overlay => new PlanImageLegendItem(overlay)).ToList().AsReadOnly();
                int structures = displayed.Count(overlay => overlay.Kind == "structure" && !hidden.Contains(overlay.Label)), dose = displayed.Count(overlay => overlay.Kind == "isodose");
                AvailabilityText = (image.Synthetic ? "Simulation" : "CT verfügbar") + " · " + structures + " Konturen · " + dose + " Isodosen";
                int unavailable = overlays.Count(overlay => overlay.SourceStatus != ReviewStatusCodes.Available);
                if (unavailable > 0) AvailabilityText += unavailable == 1 ? " · 1 Quelle fehlt" : " · " + unavailable + " Quellen fehlen";
                var viewport = PlanImageRenderer.CalculateViewport(image, focusIsocenter);
                ViewportText = string.Format(CultureInfo.CurrentCulture, "Sichtfeld {0:0} × {0:0} mm", viewport.HalfExtentMillimeters * 2);
                if (focusIsocenter && !viewport.HasIsocenter) ViewportText += " · Iso fehlt";
                if (focusIsocenter && !viewport.UsesDoseRegion) ViewportText += " · ohne Dosiszoom";
                if (viewport.UsesDoseRegion && image.DoseFocusRegion.CoverageLimited) ViewportText += " · Dosisbereich begrenzt";
                DetailsText = (image.Caption ?? "") + "\n" + (image.OverlaySummary ?? "") + "\n" +
                    (focusIsocenter ? viewport.Description : "Full CT extent, image-centered; physical pixel spacing is preserved. Outside CT coverage stays unfilled.") +
                    "\n" + string.Join("\n", overlays.Where(overlay => overlay.SourceStatus != ReviewStatusCodes.Available)
                        .Select(overlay => (overlay.Label ?? overlay.Kind) + ": " + (overlay.UnavailableReason ?? "Nicht verfügbar.")));
            }
            catch (Exception)
            {
                // Renderer boundary: expose failure without substituting CT or logging clinical source text.
                IsAvailable = false; ImageSource = null;
                AvailabilityText = "CT konnte nicht dargestellt werden. Bilddaten erneut laden.";
                LegendItems = new List<PlanImageLegendItem>().AsReadOnly();
            }
        }
        public string Kind { get; private set; }
        public string Title { get; private set; }
        public bool IsAvailable { get; private set; }
        public ImageSource ImageSource { get; private set; }
        public string AvailabilityText { get; private set; }
        public string ViewportText { get; private set; }
        public string DetailsText { get; private set; }
        public IList<PlanImageLegendItem> LegendItems { get; private set; }
    }

    public sealed class StructureVisibilityEventArgs : EventArgs
    {
        public StructureVisibilityEventArgs(string id, bool visible) { StructureId = id; IsVisible = visible; }
        public string StructureId { get; private set; }
        public bool IsVisible { get; private set; }
    }

    public sealed class PlanImageLegendItem : INotifyPropertyChanged
    {
        private bool visible = true;
        public event PropertyChangedEventHandler PropertyChanged;
        public bool CanToggle { get { return Kind == "structure"; } }
        public bool IsVisible
        {
            get { return visible; }
            set { if (visible == value) return; visible = value;
                var handler = PropertyChanged; if (handler != null) handler(this, new PropertyChangedEventArgs("IsVisible")); }
        }
        internal PlanImageLegendItem(ReviewImageOverlay overlay)
        {
            Kind = overlay.Kind; Label = overlay.Label ?? (Kind == "isodose" ? "Isodose" : "Struktur");
            ColorHex = overlay.ColorHex;
            DoseGy = overlay.DoseGy;
            try { var brush = (SolidColorBrush)new BrushConverter().ConvertFromString(ColorHex); brush.Freeze(); Color = brush; }
            catch (Exception) { ColorHex = "#FFFFFF"; Color = Brushes.White; }
            Description = (Kind == "isodose" ? "Isodose" : "Struktur") + " · " + (overlay.Source ?? "Abgelöster Bild-Snapshot");
        }
        public string Kind { get; private set; }
        public string Label { get; private set; }
        public string ColorHex { get; private set; }
        public double? DoseGy { get; private set; }
        public Brush Color { get; private set; }
        public string Description { get; private set; }
    }
}
