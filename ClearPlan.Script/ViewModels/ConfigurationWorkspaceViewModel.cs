using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using ClearPlan.Core.Configuration;

namespace ClearPlan
{
    public sealed class ConfigurationEntryViewModel
    {
        public string Key { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string Extension { get; set; }
        public Func<string> SourcePath { get; set; }
        public Action<byte[]> Validate { get; set; }
    }

    /// <summary>Detached configuration editing. Reading the page never imports or writes a file.</summary>
    public sealed class ConfigurationWorkspaceViewModel : INotifyPropertyChanged
    {
        private readonly ConfigurationHistoryStore _store;
        private readonly Action<string, string> _activateCandidate;
        private ConfigurationEntryViewModel _selectedEntry;
        private ConfigurationRevision _current;
        private ConfigurationRevision _selectedVersion;
        private string _editorText = string.Empty;
        private string _persistedText = string.Empty;
        private string _changeReason = string.Empty;
        private string _statusMessage = string.Empty;
        private string _configuredPath = string.Empty;
        private string _managedPath = string.Empty;
        private string _excelDraftPath;
        private int _excelDraftBaseRevision;
        private bool _newJsonDraft;

        public ConfigurationWorkspaceViewModel(string rootDirectory,
            IEnumerable<ConfigurationEntryViewModel> entries, Action<string, string> activateCandidate)
        {
            RootDirectory = Path.GetFullPath(rootDirectory);
            _store = new ConfigurationHistoryStore(RootDirectory);
            _activateCandidate = activateCandidate ?? ((key, path) => { });
            Entries = new ObservableCollection<ConfigurationEntryViewModel>(entries);
            Versions = new ObservableCollection<ConfigurationRevision>();
            Actor = Environment.UserDomainName + "\\" + Environment.UserName;
            Select(Entries.FirstOrDefault());
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public string RootDirectory { get; private set; }
        public string Actor { get; private set; }
        public ObservableCollection<ConfigurationEntryViewModel> Entries { get; private set; }
        public ObservableCollection<ConfigurationRevision> Versions { get; private set; }
        public ConfigurationEntryViewModel SelectedEntry { get { return _selectedEntry; } }
        public ConfigurationRevision SelectedVersion
        {
            get { return _selectedVersion; }
            set { _selectedVersion = value; Notify("SelectedVersion"); }
        }
        public int CurrentRevisionNumber { get { return _current == null ? 0 : _current.RevisionNumber; } }
        public string ConfiguredPath { get { return _configuredPath; } }
        public string ManagedPath { get { return _managedPath; } }
        public string EditingPath { get { return _excelDraftPath ?? string.Empty; } }
        public bool IsManaged { get { return _current != null; } }
        public bool IsJson { get { return _selectedEntry != null && _selectedEntry.Extension == ".json"; } }
        public bool IsExcel { get { return _selectedEntry != null && _selectedEntry.Extension == ".xlsx"; } }
        public bool CanEditJson { get { return IsJson && (IsManaged || _newJsonDraft); } }
        public bool CanEditExcel { get { return IsExcel && IsManaged; } }
        public bool CanAcceptExcel { get { return CanEditExcel && !string.IsNullOrWhiteSpace(_excelDraftPath); } }
        public bool HasUnsavedChanges { get { return _newJsonDraft || _editorText != _persistedText || _excelDraftPath != null; } }
        public string StateLabel
        {
            get
            {
                if (IsManaged) return "Verwaltet · Version " + CurrentRevisionNumber;
                return string.IsNullOrWhiteSpace(ConfiguredPath) ? "Nicht konfiguriert" :
                    File.Exists(ConfiguredPath) ? "Externe Quelle · noch nicht versioniert" : "Quelldatei fehlt oder ist nicht zugänglich";
            }
        }
        public string StatusMessage { get { return _statusMessage; } private set { _statusMessage = value; Notify("StatusMessage"); } }
        public string ChangeReason { get { return _changeReason; } set { _changeReason = value; Notify("ChangeReason"); } }
        public string EditorText
        {
            get { return _editorText; }
            set { _editorText = value ?? string.Empty; Notify("EditorText"); Notify("HasUnsavedChanges"); }
        }

        public void Select(ConfigurationEntryViewModel entry)
        {
            _selectedEntry = entry;
            ChangeReason = string.Empty;
            _excelDraftPath = null;
            _newJsonDraft = false;
            Refresh();
        }

        public void Refresh()
        {
            try
            {
                Versions.Clear();
                _current = null;
                _managedPath = string.Empty;
                _configuredPath = _selectedEntry == null || _selectedEntry.SourcePath == null
                    ? string.Empty : _selectedEntry.SourcePath() ?? string.Empty;
                if (_selectedEntry != null && Directory.Exists(Path.Combine(RootDirectory, _selectedEntry.Key)))
                {
                    foreach (var revision in _store.ListVersions(_selectedEntry.Key)) Versions.Add(revision);
                    _current = _store.GetCurrent(_selectedEntry.Key);
                    _managedPath = _store.GetManagedPath(_selectedEntry.Key) ?? string.Empty;
                }
                SelectedVersion = Versions.FirstOrDefault();
                _persistedText = IsJson && _current != null
                    ? Decode(_store.ReadVersion(_selectedEntry.Key, _current.RevisionNumber)) : string.Empty;
                EditorText = _persistedText;
                StatusMessage = IsManaged
                    ? "Versionen sind Dateistände, keine klinische Freigabe. Änderungen werden vor der Übernahme geprüft."
                    : "Zuerst eine bestehende Datei ausdrücklich als Arbeitskopie importieren. Die externe Quelle bleibt unverändert und aktiv, bis Sie die Einstellungen speichern.";
            }
            catch (Exception exception)
            {
                _current = null; _managedPath = string.Empty; _persistedText = string.Empty; EditorText = string.Empty;
                StatusMessage = "Konfiguration nicht geladen: " + exception.Message + " Pfad und Zugriffsrechte prüfen; es wurde nichts geändert.";
            }
            NotifyAll();
        }

        public bool ImportFile(string sourcePath)
        {
            return Execute(() =>
            {
                if (_selectedEntry == null) throw new InvalidOperationException("Konfiguration auswählen.");
                if (!string.Equals(Path.GetExtension(sourcePath), _selectedEntry.Extension, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Dateiformat stimmt nicht mit der gewählten Konfiguration überein.");
                byte[] content = ReadBounded(sourcePath);
                _store.Import(_selectedEntry.Key, _selectedEntry.Title, _selectedEntry.Extension,
                    content, Actor, RequireReason(), CurrentRevisionNumber, _selectedEntry.Validate);
                Complete("Datei geprüft und als neue Version importiert. Die externe Originaldatei wurde nicht verändert.");
            });
        }

        public bool SaveJson()
        {
            return Execute(() =>
            {
                if (!CanEditJson) throw new InvalidOperationException("JSON zuerst als verwaltete Arbeitskopie importieren.");
                byte[] content = new UTF8Encoding(false, true).GetBytes(EditorText);
                if (_newJsonDraft && !IsManaged)
                    _store.Import(_selectedEntry.Key, _selectedEntry.Title, ".json", content, Actor,
                        RequireReason(), 0, _selectedEntry.Validate);
                else
                    _store.Save(_selectedEntry.Key, content, Actor, RequireReason(), CurrentRevisionNumber, _selectedEntry.Validate);
                Complete("JSON geprüft und als neue Version gespeichert.");
            });
        }

        public void StageJson(string key, string json)
        {
            Select(Entries.Single(entry => entry.Key == key));
            _newJsonDraft = !IsManaged;
            EditorText = json;
            StatusMessage = "Entwurf übernommen, noch nicht gespeichert. Änderungsgrund ergänzen und als Version speichern.";
            NotifyAll();
        }

        public string PrepareExcelDraft()
        {
            string result = null;
            Execute(() =>
            {
                if (!CanEditExcel) throw new InvalidOperationException("Excel zuerst als verwaltete Arbeitskopie importieren.");
                if (_excelDraftPath == null)
                {
                    _excelDraftPath = _store.CreateTemporaryFile(".editing", ".xlsx",
                        _store.ReadVersion(_selectedEntry.Key, CurrentRevisionNumber));
                    _excelDraftBaseRevision = CurrentRevisionNumber;
                }
                result = _excelDraftPath;
                StatusMessage = "Separater Excel-Entwurf. In Excel speichern und schließen, dann „Änderung übernehmen“. Der aktive Dateistand bleibt bis dahin unverändert.";
                NotifyAll();
            });
            return result;
        }

        public bool AcceptExcelDraft()
        {
            return Execute(() =>
            {
                if (!CanAcceptExcel) throw new InvalidOperationException("Zuerst einen Excel-Entwurf öffnen.");
                _store.Save(_selectedEntry.Key, ReadBounded(_excelDraftPath), Actor,
                    RequireReason(), _excelDraftBaseRevision, _selectedEntry.Validate);
                Complete("Excel-Entwurf geprüft und als neue Version übernommen.");
            });
        }

        public bool RestoreSelected()
        {
            return Execute(() =>
            {
                if (_selectedEntry == null || SelectedVersion == null) throw new InvalidOperationException("Version zum Wiederherstellen auswählen.");
                _store.Restore(_selectedEntry.Key, SelectedVersion.RevisionNumber, Actor,
                    RequireReason(), CurrentRevisionNumber, _selectedEntry.Validate);
                Complete("Gewählter Dateistand wurde als neue Version wiederhergestellt. Die bisherige Historie bleibt erhalten.");
            });
        }

        public void CancelEdit()
        {
            _newJsonDraft = false;
            _excelDraftPath = null;
            ChangeReason = string.Empty;
            Refresh();
        }

        private string RequireReason()
        {
            if (string.IsNullOrWhiteSpace(ChangeReason)) throw new InvalidOperationException("Bitte einen Änderungsgrund eingeben.");
            return ChangeReason.Trim();
        }

        private void Complete(string message)
        {
            string path = _store.GetManagedPath(_selectedEntry.Key);
            _activateCandidate(_selectedEntry.Key, path);
            _newJsonDraft = false; _excelDraftPath = null; ChangeReason = string.Empty;
            Refresh();
            StatusMessage = message + " Versionspfad im Einstellungsentwurf aktualisiert. Zur Aktivierung immer Einstellungen speichern und Planansicht neu laden; bis dahin bleibt die bisher gespeicherte Quelle aktiv.";
        }

        private bool Execute(Action action)
        {
            try { action(); return true; }
            catch (Exception exception)
            {
                StatusMessage = "Nicht übernommen: " + exception.Message + " Entwurf prüfen; bei Versionskonflikt aktuelle Version neu laden.";
                NotifyAll();
                return false;
            }
        }

        private static byte[] ReadBounded(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (stream.Length == 0 || stream.Length > 32 * 1024 * 1024)
                    throw new InvalidOperationException("Die Datei ist leer oder größer als 32 MiB.");
                var bytes = new byte[(int)stream.Length];
                int read = 0;
                while (read < bytes.Length)
                {
                    int count = stream.Read(bytes, read, bytes.Length - read);
                    if (count == 0) throw new IOException("Die Datei wurde während des Lesens verändert.");
                    read += count;
                }
                return bytes;
            }
        }

        internal static string Decode(byte[] bytes)
        {
            string text = new UTF8Encoding(false, true).GetString(bytes);
            return text.Length > 0 && text[0] == '\uFEFF' ? text.Substring(1) : text;
        }

        private void NotifyAll() { Notify(string.Empty); }
        private void Notify(string name)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(name));
        }
    }
}
