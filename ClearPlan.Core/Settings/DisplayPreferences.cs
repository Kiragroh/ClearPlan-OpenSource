using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace ClearPlan.Core.Settings
{
    /// <summary>
    /// GUI display preferences (2026-11-20). They change only what the review window shows:
    /// the CT zoom reference isodose, its margin and the BEV layer mode. No check, value, report or
    /// log depends on them. Precedence: user file, else global file, else built-in defaults.
    /// </summary>
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class DisplayPreferences
    {
        public const string BevModeDrr = "drr";
        public const string BevModePtv = "ptv";
        public const string BevModeMlc = "mlc";

        public DisplayPreferences()
        {
            SchemaVersion = 1;
            ZoomIsodosePercent = 20;
            ZoomMarginMillimeters = 10;
            BevMode = BevModeDrr;
            CollisionFollowMlc = true;
        }

        [JsonProperty] public int SchemaVersion { get; set; }
        /// <summary>CT planes zoom to this isodose level (percent of the prescription) plus the margin.</summary>
        [JsonProperty] public double ZoomIsodosePercent { get; set; }
        [JsonProperty] public double ZoomMarginMillimeters { get; set; }
        /// <summary>drr = DRR + MLC, ptv = PTV projection + MLC, mlc = MLC only.</summary>
        [JsonProperty] public string BevMode { get; set; }
        /// <summary>Collision playback: the MLC inset follows the control point nearest to the arc angle.</summary>
        [JsonProperty] public bool CollisionFollowMlc { get; set; }

        public DisplayPreferences Copy()
        {
            return new DisplayPreferences { SchemaVersion = SchemaVersion, ZoomIsodosePercent = ZoomIsodosePercent,
                ZoomMarginMillimeters = ZoomMarginMillimeters, BevMode = BevMode, CollisionFollowMlc = CollisionFollowMlc };
        }

        /// <summary>Clamps to safe display ranges; unknown modes fall back to DRR + MLC.</summary>
        public DisplayPreferences Normalized()
        {
            var copy = Copy();
            if (double.IsNaN(copy.ZoomIsodosePercent) || double.IsInfinity(copy.ZoomIsodosePercent)) copy.ZoomIsodosePercent = 20;
            copy.ZoomIsodosePercent = Math.Max(1, Math.Min(100, copy.ZoomIsodosePercent));
            if (double.IsNaN(copy.ZoomMarginMillimeters) || double.IsInfinity(copy.ZoomMarginMillimeters)) copy.ZoomMarginMillimeters = 10;
            copy.ZoomMarginMillimeters = Math.Max(0, Math.Min(50, copy.ZoomMarginMillimeters));
            if (!new[] { BevModeDrr, BevModePtv, BevModeMlc }.Contains(copy.BevMode)) copy.BevMode = BevModeDrr;
            return copy;
        }

        private static DisplayPreferences current = new DisplayPreferences();

        /// <summary>Preferences in effect for this window. Hosts set them at start-up.</summary>
        public static DisplayPreferences Current
        {
            get { return current; }
            set
            {
                current = (value ?? new DisplayPreferences()).Normalized();
                var handler = Changed;
                if (handler != null) handler(null, EventArgs.Empty);
            }
        }

        public static event EventHandler Changed;

        /// <summary>Global/user files of the running host; null in the simulator (nothing is saved).</summary>
        public static DisplayPreferencesStore Store { get; set; }

        /// <summary>Changes the preferences for this window only; saving is a separate, explicit step.</summary>
        public static void Update(Action<DisplayPreferences> change)
        {
            var next = Current.Copy();
            if (change != null) change(next);
            Current = next;
        }
    }

    /// <summary>File locations for global and per-user display preferences.</summary>
    public sealed class DisplayPreferencesStore
    {
        public DisplayPreferencesStore(string globalPath, string userPath)
        {
            GlobalPath = globalPath;
            UserPath = userPath;
        }

        public string GlobalPath { get; private set; }
        public string UserPath { get; private set; }

        /// <summary>"user", "global" or "default": which layer the effective preferences came from.</summary>
        public string LastSource { get; private set; }

        public bool HasUserPreferences { get { return !string.IsNullOrWhiteSpace(UserPath) && File.Exists(UserPath); } }

        public DisplayPreferences Load()
        {
            var user = TryRead(UserPath);
            if (user != null) { LastSource = "user"; return user; }
            var global = TryRead(GlobalPath);
            if (global != null) { LastSource = "global"; return global; }
            LastSource = "default";
            return new DisplayPreferences();
        }

        public DisplayPreferences LoadGlobalOrDefault()
        {
            return TryRead(GlobalPath) ?? new DisplayPreferences();
        }

        public void SaveUser(DisplayPreferences preferences) { Write(UserPath, preferences); }

        public void SaveGlobal(DisplayPreferences preferences) { Write(GlobalPath, preferences); }

        /// <summary>Deletes only the user file; the global defaults apply afterwards.</summary>
        public bool ResetUser()
        {
            if (!HasUserPreferences) return false;
            File.Delete(UserPath);
            return true;
        }

        /// <summary>File name for a Windows account; only letters, digits, dot, dash and underscore.</summary>
        public static string UserFileName(string accountName)
        {
            string clean = new string((accountName ?? "").Where(c => char.IsLetterOrDigit(c) || c == '.' || c == '-' || c == '_').ToArray());
            if (string.IsNullOrWhiteSpace(clean)) clean = "unknown";
            return clean.ToLowerInvariant() + ".display.json";
        }

        private static DisplayPreferences TryRead(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;
                var value = JsonConvert.DeserializeObject<DisplayPreferences>(File.ReadAllText(path, Encoding.UTF8));
                return value == null ? null : value.Normalized();
            }
            catch (Exception) { return null; }
        }

        private static void Write(string path, DisplayPreferences preferences)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new InvalidOperationException("No preferences path configured.");
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temp = path + ".tmp";
            File.WriteAllText(temp, JsonConvert.SerializeObject((preferences ?? new DisplayPreferences()).Normalized(), Formatting.Indented), new UTF8Encoding(false));
            if (File.Exists(path)) File.Replace(temp, path, null);
            else File.Move(temp, path);
        }
    }
}
