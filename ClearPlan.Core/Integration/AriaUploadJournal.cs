using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json.Linq;

namespace ClearPlan.Core.Integration
{
    /// <summary>Protected local append-only receipts. An uncertain write blocks another attempt for the same target/plan.</summary>
    public sealed class AriaUploadJournal : IDisposable
    {
        private readonly FileStream stream;
        private readonly string patientId, planKey;
        private readonly object fileGate=new object();
        public string PreviousState { get; private set; }
        public string PreviousResourceReference { get; private set; }
        public AriaUploadJournal(string directory,string endpoint,string patientIdentifier,string activePlanKey)
        {
            if(string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(endpoint) ||
                string.IsNullOrWhiteSpace(patientIdentifier) || string.IsNullOrWhiteSpace(activePlanKey)) throw new ArgumentException("Upload receipt binding is incomplete.");
            patientId=patientIdentifier; planKey=activePlanKey;
            string key;
            using(var sha=SHA256.Create()) key=BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(
                new JArray(endpoint.TrimEnd('/'),patientIdentifier,activePlanKey).ToString(Newtonsoft.Json.Formatting.None)))).Replace("-","").ToLowerInvariant();
            Directory.CreateDirectory(directory);
            stream=new FileStream(Path.Combine(directory,key+".jsonl"),FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);
            try
            {
                if(stream.Length>1024*1024) throw new InvalidOperationException("Upload receipt history requires administrative review.");
                using(var reader=new StreamReader(stream,new UTF8Encoding(false,true),false,4096,true))
                {
                    string line;
                    while((line=reader.ReadLine())!=null)
                    {
                        if(string.IsNullOrWhiteSpace(line)) continue;
                        var row=JObject.Parse(line); PreviousState=(string)row["State"]; PreviousResourceReference=(string)row["ResourceReference"];
                    }
                }
                if(PreviousState!=null && PreviousState!="Verified" && PreviousState!="Rejected")
                    throw new InvalidOperationException("A prior ARIA upload has an uncertain result. Reconcile it in ARIA before another attempt; the protected receipt was retained.");
                stream.Position=stream.Length;
            }
            catch {stream.Dispose(); throw;}
        }
        public void Record(string state,AriaReportUploadRequest request,string resourceReference=null)
        {
            if(request==null || request.ReportPatientId!=patientId || request.ActivePlanKey!=planKey)
                throw new ArgumentException("The receipt does not match the prepared patient and plan.");
            if(state!="Pending" && state!="Uncertain" && state!="Sent" && state!="Verified" && state!="Rejected") throw new ArgumentException("Unsupported upload receipt state.");
            var row=new JObject { ["Utc"]=DateTimeOffset.UtcNow.ToString("o"), ["State"]=state,
                ["PatientReference"]=request.PatientReference,["PdfSha256"]=request.PdfSha256,["ResourceReference"]=resourceReference };
            var bytes=new UTF8Encoding(false).GetBytes(row.ToString(Newtonsoft.Json.Formatting.None)+"\n");
            lock(fileGate){stream.Write(bytes,0,bytes.Length); stream.Flush(true);}
        }
        public void Dispose(){lock(fileGate){stream.Dispose();}}
    }
}
