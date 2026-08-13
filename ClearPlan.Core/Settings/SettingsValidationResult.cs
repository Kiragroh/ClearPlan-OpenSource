using System.Collections.Generic;

namespace ClearPlan.Core.Settings
{
    public sealed class SettingsValidationResult
    {
        public SettingsValidationResult()
        {
            Errors = new List<string>();
            Warnings = new List<string>();
        }

        public List<string> Errors { get; private set; }
        public List<string> Warnings { get; private set; }

        public bool IsValid
        {
            get { return Errors.Count == 0; }
        }
    }
}
