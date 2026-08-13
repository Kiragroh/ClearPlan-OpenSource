using System;
using System.IO;
using ClearPlan.Core.Settings;

namespace ClearPlan.Core.Tests
{
    internal static class WritablePathFallbackTests
    {
        public static void UsesLocalFileWhenConfiguredParentIsUnavailable()
        {
            string root = Path.Combine(
                Path.GetTempPath(),
                "ClearPlanWritablePathTests",
                Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                string blocker = Path.Combine(root, "not-a-directory");
                File.WriteAllText(blocker, "block");
                string configured = Path.Combine(
                    blocker,
                    "ClearPlan_UserLog.csv");
                string fallback = Path.Combine(
                    root,
                    "Logs",
                    "ClearPlan_UserLog.csv");

                WritablePathResult result =
                    WritablePathFallback.EnsureFileParent(configured, fallback);

                TestAssert.True(result.UsedFallback);
                TestAssert.Equal(fallback, result.Path);
                TestAssert.True(
                    Directory.Exists(Path.GetDirectoryName(fallback)));
                TestAssert.True(
                    !string.IsNullOrWhiteSpace(result.FailureMessage));
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        public static void KeepsConfiguredDirectoryWhenItIsWritable()
        {
            string root = Path.Combine(
                Path.GetTempPath(),
                "ClearPlanWritablePathTests",
                Guid.NewGuid().ToString("N"));
            try
            {
                string configured = Path.Combine(root, "ConfiguredLogs");
                string fallback = Path.Combine(root, "FallbackLogs");

                WritablePathResult result =
                    WritablePathFallback.EnsureDirectory(configured, fallback);

                TestAssert.False(result.UsedFallback);
                TestAssert.Equal(configured, result.Path);
                TestAssert.True(Directory.Exists(configured));
                TestAssert.False(Directory.Exists(fallback));
            }
            finally
            {
                if (Directory.Exists(root))
                {
                    Directory.Delete(root, true);
                }
            }
        }
    }
}
