using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    /// <summary>Display-only levels relative to total plan prescription. Never changes dose or goals.</summary>
    public sealed class IsodoseDisplayConfiguration
    {
        public int SchemaVersion { get; set; } = 1;
        public List<IsodoseDisplayLevel> Levels { get; set; } = new List<IsodoseDisplayLevel>();

        public static IsodoseDisplayConfiguration CreateDefault()
        {
            using (var stream = typeof(IsodoseDisplayConfiguration).Assembly.GetManifestResourceStream("ClearPlan.Core.IsodoseDisplay.defaults.json"))
            using (var reader = new StreamReader(stream)) return Parse(reader.ReadToEnd());
        }
        public static IsodoseDisplayConfiguration Parse(string json)
        {
            if (string.IsNullOrWhiteSpace(json) || json.Length > 32768) throw new FormatException("Isodosen-Konfiguration ist leer oder zu groß.");
            IsodoseDisplayConfiguration config;
            try {
                var document = Newtonsoft.Json.Linq.JObject.Parse(json);
                if (document["SchemaVersion"] == null || document["Levels"] == null || document["Levels"].Type != Newtonsoft.Json.Linq.JTokenType.Array)
                    throw new FormatException("Isodosen benötigen SchemaVersion und eine Levels-Liste.");
                config = document.ToObject<IsodoseDisplayConfiguration>();
            }
            catch (JsonException) { throw new FormatException("Isodosen-Konfiguration ist kein gültiges JSON."); }
            if (config == null) throw new FormatException("Isodosen-Konfiguration fehlt.");
            config.Validate(); return config;
        }
        public static IsodoseDisplayConfiguration Load(string path, out string message)
        {
            message = "Isodosen-Defaults geladen. Anzeigeänderungen überschreiben diese Datei nicht.";
            if (string.IsNullOrWhiteSpace(path)) { message = "Standardpalette aktiv; kein eigener Default-Pfad konfiguriert."; return CreateDefault(); }
            try
            {
                // A slow optional UNC source must not freeze the native review. No ESAPI access in this task.
                var read = System.Threading.Tasks.Task.Run(() => {
                    var info = new FileInfo(path);
                    if (!info.Exists || info.Length > 32768) throw new FormatException("Missing or oversized display configuration.");
                    return Parse(File.ReadAllText(path));
                });
                if (read.Wait(TimeSpan.FromSeconds(2))) return read.Result;
                read.ContinueWith(t => { var observed = t.Exception; }, System.Threading.Tasks.TaskContinuationOptions.OnlyOnFaulted);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is AggregateException || ex is ArgumentException) { }
            message = "Isodosen-Defaults nicht lesbar oder ungültig: Standardpalette aktiv. Pfad und Konfiguration in Settings prüfen.";
            return CreateDefault();
        }
        public void Validate()
        {
            if (SchemaVersion != 1 || Levels == null || Levels.Count > 24)
                throw new FormatException("Isodosen: Schema 1 und maximal 24 Stufen erforderlich.");
            if (Levels.Any(l => l == null || !Finite(l.Percent) || l.Percent <= 0 || l.Percent > 1000 ||
                l.ColorHex == null || !Regex.IsMatch(l.ColorHex, "^#[0-9a-fA-F]{6}$")))
                throw new FormatException("Jede Isodose benötigt einen Wert > 0 bis 1000 % Rx und eine Farbe #RRGGBB.");
            if (Levels.GroupBy(l => Math.Round(l.Percent, 6)).Any(g => g.Count() != 1))
                throw new FormatException("Isodosenstufen dürfen nicht doppelt vorkommen.");
        }
        public IsodoseDisplayConfiguration Copy() { return Parse(JsonConvert.SerializeObject(this)); }

        public ReviewPlanImage Apply(ReviewPlanImage source)
        {
            Validate();
            if (source == null) throw new ArgumentNullException("source");
            var image = JsonConvert.DeserializeObject<ReviewPlanImage>(JsonConvert.SerializeObject(source));
            image.Overlays = image.Overlays ?? new List<ReviewImageOverlay>();
            if (source.DosePlane == null)
            {
                // Older captures contain contours only. Recolor existing levels, never invent missing dose samples.
                image.Overlays = image.Overlays.Where(o => o == null || o.Kind != "isodose" ||
                    !PercentFor(o).HasValue || Levels.Any(l => l.Enabled && Math.Abs(l.Percent - PercentFor(o).Value) < 0.000001)).ToList();
                foreach (var overlay in image.Overlays.Where(o => o != null && o.Kind == "isodose" && PercentFor(o).HasValue))
                    overlay.ColorHex = Levels.Single(l => l.Enabled && Math.Abs(l.Percent - PercentFor(overlay).Value) < 0.000001).ColorHex;
                return image;
            }
            var existingDose = image.Overlays.Where(o => o != null && o.Kind == "isodose" && o.SourceStatus == ReviewStatusCodes.Available).ToList();
            image.Overlays.RemoveAll(o => o != null && o.Kind == "isodose");
            var plane = source.DosePlane;
            if (!plane.IsValid || image.WidthPixels < 2 || image.HeightPixels < 2)
            {
                image.Overlays.Add(new ReviewImageOverlay { Kind = "isodose", Label = "Isodoses", ColorHex = "#AAB4C0",
                    SourceStatus = ReviewStatusCodes.Unavailable, UnavailableReason = "Detached dose plane or prescription is invalid; no substitute dose." });
                return image;
            }
            double maximum = plane.SamplesGy.Where(Finite).DefaultIfEmpty(double.NaN).Max();
            foreach (var level in Levels.Where(l => l.Enabled).OrderBy(l => l.Percent))
            {
                double gy = plane.PrescriptionGy * level.Percent / 100.0;
                if (!Finite(maximum) || gy > maximum) continue;
                var existing = existingDose.FirstOrDefault(o => o.DoseGy.HasValue && Math.Abs(o.DoseGy.Value - gy) < 1e-9 && o.Paths != null && o.Paths.Count > 0);
                List<ReviewImagePath> paths;
                try { paths = existing != null ? existing.Paths : PlanImageOverlayGeometry.TraceIsodose(plane.SamplesGy, plane.Columns, plane.Rows,
                    gy, image.WidthPixels, image.HeightPixels, 20000); }
                catch (ArgumentException) { image.Overlays.Add(new ReviewImageOverlay { Kind = "isodose", Label = level.Percent.ToString("0.###", CultureInfo.InvariantCulture) + "% Rx",
                    SourceStatus = ReviewStatusCodes.Unavailable, ColorHex = level.ColorHex, UnavailableReason = "Display contour complexity exceeded; level not rendered." }); continue; }
                if (paths.Count == 0) continue;
                image.Overlays.Add(new ReviewImageOverlay { Kind = "isodose", DoseGy = gy, PrescriptionPercent = level.Percent,
                    Label = string.Format(CultureInfo.InvariantCulture, "{0:0.###}% Rx / {1:0.##} Gy", level.Percent, gy),
                    ColorHex = level.ColorHex.ToUpperInvariant(), Source = plane.Source,
                    SourceStatus = ReviewStatusCodes.Available, Paths = paths });
            }
            return image;
        }
        public static double? PercentFor(ReviewImageOverlay overlay)
        {
            if (overlay == null) return null;
            if (overlay.PrescriptionPercent.HasValue && Finite(overlay.PrescriptionPercent.Value)) return overlay.PrescriptionPercent;
            var match = Regex.Match(overlay.Label ?? "", "^([0-9]+(?:\\.[0-9]+)?)% Rx(?: /|$)");
            double value;
            return match.Success && double.TryParse(match.Groups[1].Value, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out value) ? (double?)value : null;
        }
        internal static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
    }
    public sealed class IsodoseDisplayLevel
    {
        public double Percent { get; set; }
        public string ColorHex { get; set; }
        public bool Enabled { get; set; } = true;
    }
}
