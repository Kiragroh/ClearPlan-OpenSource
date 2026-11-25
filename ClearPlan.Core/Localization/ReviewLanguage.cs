using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace ClearPlan.Core.Localization
{
    /// <summary>Presentation only. Never apply to stored clinical identifiers or values.</summary>
    public static class ReviewLanguage
    {
        private static string code = "de";
        [ThreadStatic] private static string renderCode;
        private static readonly Lazy<Dictionary<string,string>> translations = new Lazy<Dictionary<string,string>>(LoadCatalog);
        public static event EventHandler Changed;
        public static string Code
        {
            get { return renderCode ?? code; }
            set {
                string next = Normalize(value);
                if (code == next) return;
                code = next;
                var handler = Changed; if (handler != null) handler(null,EventArgs.Empty);
            }
        }
        public static string Normalize(string value)
        {
            if (string.Equals(value,"en",StringComparison.OrdinalIgnoreCase) || string.Equals(value,"en-US",StringComparison.OrdinalIgnoreCase)) return "en";
            if (string.Equals(value,"de",StringComparison.OrdinalIgnoreCase) || string.Equals(value,"de-DE",StringComparison.OrdinalIgnoreCase)) return "de";
            throw new ArgumentException("Supported review languages: de, en.","value");
        }
        public static string Label(string german, string english)
        { return Code == "en" ? english : german; }

        public static string Text(string source)
        {
            if (source == null || Code != "en") return source;
            string translated;
            return translations.Value.TryGetValue(source,out translated) ? translated : source;
        }
        // For explicitly designated generated display messages, never identity cells.
        public static string Display(string source)
        {
            if (string.IsNullOrEmpty(source) || Code != "en") return source;
            string exact = Text(source); if (exact != source) return exact;
            // Generated values are combined from catalog phrases. Only whole literal
            // fragments are translated; numbers, units and placeholders remain intact.
            return displayPattern.Value.Replace(source, m => translations.Value[m.Value]);
        }
        private static readonly Lazy<Regex> displayPattern = new Lazy<Regex>(() => new Regex(
            string.Join("|", translations.Value.Keys.Where(k => k.Length >= 5 && !k.Contains("{") && !k.Contains("\\"))
              .OrderByDescending(k=>k.Length).Select(k => "(?<![\\p{L}\\p{N}_])" + Regex.Escape(k) + "(?![\\p{L}\\p{N}_])")), RegexOptions.CultureInvariant));
        public static string Format(string template, params object[] values)
        { return string.Format(CultureInfo.InvariantCulture,Text(template),values); }
        public static IDisposable Scope(string language) { return new RenderScope(Normalize(language)); }
        private sealed class RenderScope : IDisposable
        {
            private readonly string previous;
            public RenderScope(string language) { previous=renderCode;renderCode=language; }
            public void Dispose() { renderCode=previous; }
        }
        private static Dictionary<string,string> LoadCatalog()
        {
            using(var stream=typeof(ReviewLanguage).Assembly.GetManifestResourceStream("ClearPlan.ReviewLanguage.en.json"))
            using(var reader=new StreamReader(stream,Encoding.UTF8))
                return JsonConvert.DeserializeObject<Dictionary<string,string>>(reader.ReadToEnd());
        }
        public static string PreferencePath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"ClearPlan","review-language.txt"); } }
        public static void LoadPreference()
        {
            string value=Environment.GetEnvironmentVariable("CLEARPLAN_LANGUAGE");
            try { if (string.IsNullOrWhiteSpace(value) && File.Exists(PreferencePath)) value=File.ReadAllText(PreferencePath).Trim(); }
            catch(IOException) { } catch(UnauthorizedAccessException) { }
            if (string.IsNullOrWhiteSpace(value)) return;
            try { Code=Normalize(value); } catch(ArgumentException) { }
        }
        public static bool SavePreference()
        {
            try { Directory.CreateDirectory(Path.GetDirectoryName(PreferencePath)); File.WriteAllText(PreferencePath,code,Encoding.UTF8); return true; }
            catch(IOException) { return false; } catch(UnauthorizedAccessException) { return false; }
        }
    }
}
