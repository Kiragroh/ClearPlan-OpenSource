using System.IO;

namespace ClearPlan.Core.Settings
{
    public static class SettingsPathResolver
    {
        public static string Resolve(string baseDirectory, string relativeOrAbsolutePath)
        {
            string resolvedBase = Path.GetFullPath(
                string.IsNullOrWhiteSpace(baseDirectory) ? "." : baseDirectory);
            if (string.IsNullOrWhiteSpace(relativeOrAbsolutePath))
            {
                return resolvedBase;
            }

            return Path.IsPathRooted(relativeOrAbsolutePath)
                ? Path.GetFullPath(relativeOrAbsolutePath)
                : Path.GetFullPath(Path.Combine(resolvedBase, relativeOrAbsolutePath));
        }
    }
}
