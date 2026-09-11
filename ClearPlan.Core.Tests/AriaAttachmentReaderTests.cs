using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ClearPlan.Core.Integration;
using Newtonsoft.Json.Linq;

namespace ClearPlan.Core.Tests
{
    internal static class AriaAttachmentReaderTests
    {
        public static void ConfigurationAndPathBoundary()
        {
            var property = typeof(AriaUploadConfiguration).GetProperty("AttachmentReadbackRoots");
            TestAssert.NotNull(property, "Attachment filesystem readback must be disabled until explicit roots are configured.");
            TestAssert.Equal(0, ((string[])property.GetValue(AriaUploadConfiguration.Parse("{\"Enabled\":false}"))).Length);
            object reader = Reader("\\\\synthetic.invalid\\share\\approved");
            TestAssert.True(Allowed(reader, "\\\\SYNTHETIC.INVALID\\SHARE\\approved\\folder\\report.pdf"));
            foreach (string path in new[] { "\\\\synthetic.invalid\\share\\approved-other\\report.pdf", "\\\\other.invalid\\share\\approved\\report.pdf",
                "\\\\synthetic.invalid\\share\\approved\\..\\private\\report.pdf", "\\\\synthetic.invalid\\share\\approved\\report.pdf:stream",
                "https://synthetic.invalid/share/approved/report.pdf", "file://synthetic.invalid/share/approved/report.pdf", "C:\\approved\\report.pdf",
                "\\\\?\\UNC\\synthetic.invalid\\share\\approved\\report.pdf", "\\\\synthetic.invalid\\share\\approved\\folder.\\report.pdf",
                "\\\\synthetic.invalid\\share\\approved\\folder \\report.pdf", "\\\\synthetic.invalid\\share\\approved", "\\\\synthetic.invalid\\share\\approved\\\\report.pdf",
                "\\\\synthetic.invalid\\share\\approved\\NUL.pdf", "\\\\synthetic.invalid\\share\\approved\\COM1\\report.pdf",
                "\\\\synthetic.invalid\\share\\approved\\" + new string('x', 300) + ".pdf" })
                TestAssert.False(Allowed(reader, path), "Only an unambiguous descendant of an approved UNC root is readable.");
            TestAssert.False(Allowed(Reader(), "\\\\synthetic.invalid\\share\\approved\\report.pdf"));
            foreach (string root in new[] { "C:\\private", "https://synthetic.invalid/path", "\\\\synthetic.invalid", "\\\\synthetic.invalid\\share\\..\\other", "\\\\?\\UNC\\server\\share" })
            {
                var json = new JObject { ["Enabled"] = false, ["AttachmentReadbackRoots"] = new JArray(root) };
                TestAssert.Throws<FormatException>(() => AriaUploadConfiguration.Parse(json.ToString()));
            }
            var rootsJson = new JObject { ["Enabled"] = false, ["AttachmentReadbackRoots"] = new JArray(Enumerable.Range(0, 9).Select(i => "\\\\synthetic.invalid\\share\\r" + i)) };
            TestAssert.Throws<FormatException>(() => AriaUploadConfiguration.Parse(rootsJson.ToString()));
            string source = File.ReadAllText(Path.Combine("ClearPlan.Script", "MainView.AriaUpload.cs"));
            TestAssert.True(source.Contains("new AriaAttachmentReader(config.AttachmentReadbackRoots)") && source.Contains("attachmentReader.ReadAsync"), "Native uploads must use only configured attachment roots.");
            TestAssert.True(source.Contains("verification.AttachmentCreationChecked"), "Omitted attachment creation must remain visible as unchecked.");
            TestAssert.True(source.Contains("string.IsNullOrEmpty(result.ResourceReference)"), "A validated create reference permits GET-only recovery of an uncertain response.");
            var example = JObject.Parse(File.ReadAllText(Path.Combine("ClearPlan.Script", "Distribution", "AriaUpload.example.json")));
            TestAssert.True(example["AttachmentReadbackRoots"] is JArray && !example["AttachmentReadbackRoots"].Any(), "Public file readback must remain explicitly disabled.");
        }

        public static void BlockBeforeFileAccess()
        {
            object reader = Reader("\\\\synthetic.invalid\\share\\approved");
            var method = reader.GetType().GetMethod("ReadAsync");
            var rejected = (Task<byte[]>)method.Invoke(reader, new object[] { "\\\\outside.invalid\\share\\report.pdf", CancellationToken.None });
            TestAssert.Throws<InvalidOperationException>(() => rejected.GetAwaiter().GetResult());
            using (var cancel = new CancellationTokenSource())
            {
                cancel.Cancel();
                var cancelled = (Task<byte[]>)method.Invoke(reader, new object[] { "\\\\synthetic.invalid\\share\\approved\\report.pdf", cancel.Token });
                TestAssert.Throws<OperationCanceledException>(() => cancelled.GetAwaiter().GetResult());
            }
        }

        public static void LocalFileAndReparseGuards()
        {
            var type = ReaderType();
            var read = type.GetMethod("ReadFileLocked", BindingFlags.Static | BindingFlags.NonPublic);
            var attributes = type.GetMethod("ValidateAttributes", BindingFlags.Static | BindingFlags.NonPublic);
            TestAssert.NotNull(read); TestAssert.NotNull(attributes);
            foreach (uint flags in new uint[] { 0x400, 0x410 })
                RejectInvocation(() => attributes.Invoke(null, new object[] { flags, (flags & 0x10) != 0 }));
            string directory = Path.Combine(Path.GetTempPath(), "ClearPlanAttachmentTest-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                string file = Path.Combine(directory, "synthetic.pdf");
                byte[] bytes = Encoding.ASCII.GetBytes("%PDF-synthetic-read-only"); File.WriteAllBytes(file, bytes);
                var actual = (byte[])read.Invoke(null, new object[] { file, CancellationToken.None });
                TestAssert.True(bytes.SequenceEqual(actual));
                TestAssert.True(bytes.SequenceEqual(File.ReadAllBytes(file)), "Reading must not modify the stored PDF.");
                using (var stream = new FileStream(file, FileMode.Open, FileAccess.Write, FileShare.None)) stream.SetLength(AriaReportUploadRequest.MaximumPdfBytes + 1L);
                RejectInvocation(() => read.Invoke(null, new object[] { file, CancellationToken.None }));
                File.WriteAllText(file, "not a PDF");
                RejectInvocation(() => read.Invoke(null, new object[] { file, CancellationToken.None }));
            }
            finally
            {
                if (Path.GetDirectoryName(directory) == Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar) && Path.GetFileName(directory).StartsWith("ClearPlanAttachmentTest-", StringComparison.Ordinal))
                    Directory.Delete(directory, true);
            }
        }

        private static Type ReaderType()
        { var type = typeof(AriaReportUploadRequest).Assembly.GetType("ClearPlan.Core.Integration.AriaAttachmentReader"); TestAssert.NotNull(type, "A constrained file reader is required."); return type; }
        private static object Reader(params string[] roots) { return Activator.CreateInstance(ReaderType(), new object[] { roots }); }
        private static bool Allowed(object reader, string path) { return (bool)reader.GetType().GetMethod("IsAllowedPath").Invoke(reader, new object[] { path }); }
        private static void RejectInvocation(Action action)
        { var error = TestAssert.Throws<TargetInvocationException>(action); TestAssert.True(error.InnerException is InvalidOperationException); }
    }
}
