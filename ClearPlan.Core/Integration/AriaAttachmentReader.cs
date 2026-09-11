using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32.SafeHandles;

namespace ClearPlan.Core.Integration
{
    /// <summary>Read-only PDF verification under explicitly configured UNC roots. Never opens URLs or shell handlers.</summary>
    public sealed class AriaAttachmentReader
    {
        private readonly string[] roots;
        public AriaAttachmentReader(string[] allowedRoots) { roots = ValidateRoots(allowedRoots); }

        internal static string[] ValidateRoots(string[] values)
        {
            if (values == null || values.Length > 8) throw new ArgumentException("Attachment readback roots are invalid.");
            return values.Select(value => NormalizeUnc(value, true)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        }

        public bool IsAllowedPath(string value)
        {
            try
            {
                string path = NormalizeUnc(value, false);
                return roots.Any(root => path.StartsWith(root + "\\", StringComparison.OrdinalIgnoreCase));
            }
            catch (ArgumentException) { return false; }
        }

        public async Task<byte[]> ReadAsync(string path, CancellationToken cancellation)
        {
            cancellation.ThrowIfCancellationRequested();
            if (!IsAllowedPath(path)) throw new InvalidOperationException("The attachment is not under an explicitly configured UNC readback root.");
            string normalized = NormalizeUnc(path, false);
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellation))
            {
                timeout.CancelAfter(TimeSpan.FromSeconds(10));
                try { return await AriaFileWork.RunAsync(() => ReadFileLocked(normalized, timeout.Token), timeout.Token).ConfigureAwait(false); }
                catch (OperationCanceledException)
                {
                    if (cancellation.IsCancellationRequested) throw new OperationCanceledException(cancellation);
                    throw new TimeoutException("The configured attachment readback timed out.");
                }
                catch (Exception) { throw new InvalidOperationException("The configured attachment cannot be read safely for PDF verification."); }
            }
        }

        private static string NormalizeUnc(string value, bool root)
        {
            if (string.IsNullOrWhiteSpace(value) || value != value.Trim() || value.Length > 1024 || !value.StartsWith("\\\\", StringComparison.Ordinal) ||
                value.StartsWith("\\\\?\\", StringComparison.Ordinal) || value.StartsWith("\\\\.\\", StringComparison.Ordinal) ||
                Regex.IsMatch(value, @"[\p{Cc}\p{Cf}\p{Zl}\p{Zp}/:<>""|?*]"))
                throw new ArgumentException("An unambiguous Windows UNC path is required.");
            string normalized = root ? value.TrimEnd('\\') : value;
            string[] parts = normalized.Substring(2).Split('\\');
            if (parts.Length < (root ? 2 : 3) || parts.Any(part => string.IsNullOrEmpty(part) || part.Length > 255 || part == "." || part == ".." ||
                part.EndsWith(".", StringComparison.Ordinal) || part.EndsWith(" ", StringComparison.Ordinal)))
                throw new ArgumentException("An unambiguous Windows UNC path is required.");
            if (parts.Skip(2).Any(part => Regex.IsMatch(part, @"^(CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³]|CONIN\$|CONOUT\$)(\.|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)))
                throw new ArgumentException("Windows device names are not permitted in attachment paths.");
            string fullPath;
            try { fullPath = Path.GetFullPath(normalized); }
            catch (Exception error) when (error is ArgumentException || error is IOException || error is NotSupportedException || error is System.Security.SecurityException)
            { throw new ArgumentException("The attachment path cannot be normalized safely."); }
            if (!string.Equals(fullPath, normalized, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Attachment paths must not require normalization or traversal.");
            return normalized;
        }

        private static byte[] ReadFileLocked(string path, CancellationToken cancellation)
        {
            var parents = new List<SafeFileHandle>();
            try
            {
                cancellation.ThrowIfCancellationRequested();
                string directory = Path.GetDirectoryName(path);
                string current = Path.GetPathRoot(path);
                var parentPaths = new List<string> { current };
                foreach (string segment in directory.Substring(current.Length).Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries))
                { current = Path.Combine(current, segment); parentPaths.Add(current); }
                foreach (string parent in parentPaths)
                {
                    cancellation.ThrowIfCancellationRequested();
                    // Hold each ancestor without FILE_SHARE_DELETE so a checked directory cannot be renamed underneath the file open.
                    SafeFileHandle handle = CreateFile(parent, 0x80, 3, IntPtr.Zero, 3, 0x02000000 | 0x00200000, IntPtr.Zero);
                    parents.Add(handle);
                    if (handle.IsInvalid) throw new InvalidOperationException("An attachment directory cannot be opened safely.");
                    FileInformation info;
                    if (!GetFileInformationByHandle(handle, out info)) throw new InvalidOperationException("Attachment directory attributes are unavailable.");
                    ValidateAttributes(info.Attributes, true);
                }
                cancellation.ThrowIfCancellationRequested();
                using (var handle = CreateFile(path, 0x80000000, 1, IntPtr.Zero, 3, 0x00200000 | 0x08000000, IntPtr.Zero))
                {
                    if (handle.IsInvalid) throw new InvalidOperationException("The attachment cannot be opened safely.");
                    FileInformation info;
                    if (!GetFileInformationByHandle(handle, out info)) throw new InvalidOperationException("Attachment attributes are unavailable.");
                    ValidateAttributes(info.Attributes, false);
                    ulong length = ((ulong)info.FileSizeHigh << 32) | info.FileSizeLow;
                    if (length < 5 || length > AriaReportUploadRequest.MaximumPdfBytes) throw new InvalidOperationException("The attachment PDF exceeds its permitted size.");
                    using (var input = new FileStream(handle, FileAccess.Read, 8192, false))
                    using (var output = new MemoryStream((int)length))
                    {
                        var buffer = new byte[8192]; int count;
                        while (true)
                        {
                            cancellation.ThrowIfCancellationRequested();
                            count = input.Read(buffer, 0, buffer.Length);
                            if (count == 0) break;
                            if (output.Length + count > AriaReportUploadRequest.MaximumPdfBytes) throw new InvalidOperationException("The attachment PDF exceeds its permitted size.");
                            output.Write(buffer, 0, count);
                        }
                        cancellation.ThrowIfCancellationRequested();
                        byte[] bytes = output.ToArray();
                        if ((ulong)bytes.Length != length || bytes[0] != '%' || bytes[1] != 'P' || bytes[2] != 'D' || bytes[3] != 'F' || bytes[4] != '-')
                            throw new InvalidOperationException("The stored attachment is not a stable PDF document.");
                        return bytes;
                    }
                }
            }
            finally { for (int index = parents.Count - 1; index >= 0; index--) parents[index].Dispose(); }
        }

        private static void ValidateAttributes(uint attributes, bool directory)
        {
            if ((attributes & 0x400) != 0 || ((attributes & 0x10) != 0) != directory)
                throw new InvalidOperationException("Reparse points and unexpected attachment filesystem objects are not allowed.");
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FileInformation
        {
            public uint Attributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME CreationTime, LastAccessTime, LastWriteTime;
            public uint VolumeSerialNumber, FileSizeHigh, FileSizeLow, NumberOfLinks, FileIndexHigh, FileIndexLow;
        }
        [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string fileName, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandle(SafeFileHandle file, out FileInformation information);
    }
}
