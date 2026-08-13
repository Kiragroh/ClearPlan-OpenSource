using System;
using System.IO;
using System.Security;

namespace ClearPlan.Core.Settings
{
    public sealed class WritablePathResult
    {
        public string Path { get; set; }
        public bool UsedFallback { get; set; }
        public string FailureMessage { get; set; }
    }

    public static class WritablePathFallback
    {
        public static WritablePathResult EnsureDirectory(
            string configuredDirectory,
            string fallbackDirectory)
        {
            return Ensure(
                configuredDirectory,
                fallbackDirectory,
                path =>
                {
                    Directory.CreateDirectory(path);
                });
        }

        public static WritablePathResult EnsureFileParent(
            string configuredFile,
            string fallbackFile)
        {
            return Ensure(
                configuredFile,
                fallbackFile,
                path =>
                {
                    string directory = System.IO.Path.GetDirectoryName(path);
                    if (!string.IsNullOrWhiteSpace(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }
                });
        }

        private static WritablePathResult Ensure(
            string configuredPath,
            string fallbackPath,
            Action<string> ensureWritableParent)
        {
            try
            {
                ensureWritableParent(configuredPath);
                return new WritablePathResult
                {
                    Path = configuredPath,
                    UsedFallback = false,
                    FailureMessage = string.Empty
                };
            }
            catch (Exception exception) when (IsPathFailure(exception))
            {
                ensureWritableParent(fallbackPath);
                return new WritablePathResult
                {
                    Path = fallbackPath,
                    UsedFallback = true,
                    FailureMessage = exception.Message
                };
            }
        }

        private static bool IsPathFailure(Exception exception)
        {
            return exception is IOException ||
                   exception is UnauthorizedAccessException ||
                   exception is ArgumentException ||
                   exception is NotSupportedException ||
                   exception is PathTooLongException ||
                   exception is SecurityException;
        }
    }
}
