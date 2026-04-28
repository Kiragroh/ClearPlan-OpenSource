using System;
using System.IO;
using System.Linq;

namespace ClearPlan.Helpers
{
    public static class DocumentationHelper
    {
        public static string ReadTextOrDefault(string filePath, string fallbackText)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    return File.ReadAllText(filePath);
                }
            }
            catch
            {
            }

            return fallbackText;
        }

        public static string GetLatestSection(string filePath, string version)
        {
            string content = ReadTextOrDefault(filePath, string.Empty);
            if (string.IsNullOrWhiteSpace(content))
            {
                return string.Empty;
            }

            string normalized = content.Replace("\r\n", "\n");
            string[] lines = normalized.Split('\n');

            int startIndex = FindSectionStart(lines, version);
            if (startIndex < 0)
            {
                return content.Trim();
            }

            int endIndex = FindNextSectionStart(lines, startIndex + 1);
            int length = (endIndex < 0 ? lines.Length : endIndex) - startIndex;

            return string.Join(Environment.NewLine, lines.Skip(startIndex).Take(length)).Trim();
        }

        private static int FindSectionStart(string[] lines, string version)
        {
            if (!string.IsNullOrWhiteSpace(version))
            {
                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].StartsWith("## ", StringComparison.Ordinal) &&
                        lines[i].IndexOf(version, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return i;
                    }
                }
            }

            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].StartsWith("## ", StringComparison.Ordinal))
                {
                    return i;
                }
            }

            return -1;
        }

        private static int FindNextSectionStart(string[] lines, int startIndex)
        {
            for (int i = startIndex; i < lines.Length; i++)
            {
                if (lines[i].StartsWith("## ", StringComparison.Ordinal))
                {
                    return i;
                }
            }

            return -1;
        }
    }
}
