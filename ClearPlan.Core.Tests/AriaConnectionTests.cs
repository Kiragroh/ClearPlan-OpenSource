using System;
using System.Reflection;
using System.IO;
using System.Net.Security;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using Newtonsoft.Json.Linq;
using ClearPlan.Core.Integration;

namespace ClearPlan.Core.Tests
{
    internal static class AriaConnectionTests
    {
        public static void ConfigurationContract()
        {
            var type=typeof(AriaReportDocumentBuilder).Assembly.GetType("ClearPlan.Core.Integration.AriaUploadConfiguration");
            TestAssert.True(type!=null, "ARIA connection must use an explicit external configuration, not embedded site credentials.");
            var parse=type.GetMethod("Parse", BindingFlags.Public|BindingFlags.Static);
            TestAssert.True(parse!=null);
            var disabled=parse.Invoke(null,new object[]{"{\"SchemaVersion\":1,\"Enabled\":false}"});
            TestAssert.Equal(false,(bool)type.GetProperty("Enabled").GetValue(disabled));
            var configured=AriaUploadConfiguration.Parse(ValidJson().ToString());
            TestAssert.True(configured.Enabled);
            TestAssert.Equal("PDF",configured.CategoryCode);
            TestAssert.True(string.IsNullOrEmpty(configured.FhirCertificateSha256),"Public connections default to normal certificate validation.");
        }

        public static void InvalidConfigurations()
        {
            foreach(var change in new Action<JObject>[] {
                j=>j["SchemaVersion"]=2, j=>j["TimeoutSeconds"]=121, j=>j["TimeoutSeconds"]=0,
                j=>j["BaseUrl"]="http://example.invalid/fhir/r4", j=>j["TokenUrl"]="https://u:p@example.invalid/token",
                j=>j["BaseUrl"]="https://example.invalid/fhir/r4?secret=bad", j=>j["TokenUrl"]="https://example.invalid/token#fragment",
                j=>j["FhirCertificateSha256"]="*",j=>j["FhirCertificateSha256"]=new string('z',64),
                j=>j["ClientSecret"]="SYNTHETIC_SECRET_MUST_NOT_ENTER_CONFIG",j=>j["VerifyTls"]=false,
                j=>j["ProviderReference"]="Organization/../other",j=>j["CredentialEnvironmentVariable"]="BAD\r\nKEY",
                j=>j["CategoryCode"]="Patient Document",j=>j["DocumentTypeDisplay"]="",j=>j["ClientId"]="",
                j=>j["CredentialEnvFilePath"]="https://example.invalid/secret.env"
            })
            {
                var json=ValidJson(); change(json);
                Reject(()=>AriaUploadConfiguration.Parse(json.ToString()));
            }
            Reject(()=>AriaUploadConfiguration.Parse("{\"SchemaVersion\":1,\"Enabled\":false,\"Unknown\":true}"));
            Reject(()=>AriaUploadConfiguration.Parse(new string(' ',32769)));
        }

        public static void ExternalCredentials()
        {
            string variable="CLEARPLAN_TEST_SECRET_"+Guid.NewGuid().ToString("N");
            string directory=Path.Combine(Path.GetTempPath(),"ClearPlanAriaTest-"+Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                var json=ValidJson(); json["CredentialEnvironmentVariable"]=variable; json["CredentialEnvFilePath"]="secret.env";
                var config=AriaUploadConfiguration.Parse(json.ToString());
                string file=Path.Combine(directory,"secret.env");
                File.WriteAllText(file,"# no executable dotenv processing\n"+variable+"=\"SYNTHETIC=a#b!\"\n");
                string configFile=Path.Combine(directory,"aria.json");
                TestAssert.Equal("SYNTHETIC=a#b!",config.LoadSecretAsync(configFile,CancellationToken.None).GetAwaiter().GetResult());
                Environment.SetEnvironmentVariable(variable,"SYNTHETIC_ENV_PRECEDENCE");
                TestAssert.Equal("SYNTHETIC_ENV_PRECEDENCE",config.LoadSecretAsync(configFile,CancellationToken.None).GetAwaiter().GetResult());
                Environment.SetEnvironmentVariable(variable,null);
                File.WriteAllText(file,variable+"=a\n"+variable+"=b");
                Reject(()=>config.LoadSecretAsync(configFile,CancellationToken.None).GetAwaiter().GetResult());
                File.WriteAllText(file,new string('a',8193));
                Reject(()=>config.LoadSecretAsync(configFile,CancellationToken.None).GetAwaiter().GetResult());
                using(var stop=new CancellationTokenSource())
                {
                    stop.Cancel(); bool canceled=false;
                    try {config.LoadSecretAsync(configFile,stop.Token).GetAwaiter().GetResult();}catch(OperationCanceledException){canceled=true;}
                    TestAssert.True(canceled);
                }
                File.WriteAllText(configFile,json.ToString());
                TestAssert.True(AriaUploadConfiguration.LoadAsync(configFile,CancellationToken.None).GetAwaiter().GetResult().Enabled);
            }
            finally
            {
                Environment.SetEnvironmentVariable(variable,null);
                if(Path.GetDirectoryName(directory)==Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar) &&
                    Path.GetFileName(directory).StartsWith("ClearPlanAriaTest-",StringComparison.Ordinal)) Directory.Delete(directory,true);
            }
        }

        public static void CertificateBinding()
        {
            var now=DateTime.UtcNow; var endpoint=new Uri("https://example.invalid:55370/fhir/r4");
            using(var rsa=RSA.Create())
            {
                rsa.KeySize=2048;
                var request=new CertificateRequest("CN=example.invalid",rsa,HashAlgorithmName.SHA256,RSASignaturePadding.Pkcs1);
                using(var certificate=request.CreateSelfSigned(now.AddDays(-1),now.AddDays(1)))
                using(var sha=SHA256.Create())
                {
                    string pin=BitConverter.ToString(sha.ComputeHash(certificate.RawData)).Replace("-","");
                    var unknown=new[]{X509ChainStatusFlags.UntrustedRoot};
                    TestAssert.True(AriaCertificatePolicy.Validate(endpoint,endpoint,null,certificate,SslPolicyErrors.None,new X509ChainStatusFlags[0],now));
                    TestAssert.False(AriaCertificatePolicy.Validate(endpoint,endpoint,null,certificate,SslPolicyErrors.RemoteCertificateChainErrors,unknown,now));
                    TestAssert.True(AriaCertificatePolicy.Validate(endpoint,endpoint,pin,certificate,SslPolicyErrors.RemoteCertificateChainErrors,unknown,now));
                    TestAssert.False(AriaCertificatePolicy.Validate(endpoint,endpoint,new string('0',64),certificate,SslPolicyErrors.None,unknown,now));
                    foreach(var error in new[]{SslPolicyErrors.RemoteCertificateNameMismatch,SslPolicyErrors.RemoteCertificateNotAvailable})
                        TestAssert.False(AriaCertificatePolicy.Validate(endpoint,endpoint,pin,certificate,error,unknown,now));
                    foreach(var error in new[]{X509ChainStatusFlags.Revoked,X509ChainStatusFlags.NotTimeValid,X509ChainStatusFlags.RevocationStatusUnknown,X509ChainStatusFlags.PartialChain})
                        TestAssert.False(AriaCertificatePolicy.Validate(endpoint,endpoint,pin,certificate,SslPolicyErrors.RemoteCertificateChainErrors,new[]{error},now));
                    TestAssert.False(AriaCertificatePolicy.Validate(endpoint,new Uri("https://other.invalid:55370/fhir/r4"),pin,certificate,SslPolicyErrors.None,unknown,now));
                    TestAssert.False(AriaCertificatePolicy.Validate(endpoint,new Uri("https://example.invalid:443/fhir/r4"),pin,certificate,SslPolicyErrors.None,unknown,now));
                    TestAssert.False(AriaCertificatePolicy.Validate(endpoint,endpoint,pin,certificate,SslPolicyErrors.None,unknown,now.AddDays(2)));
                }
            }
        }

        public static void DurableUploadReceipts()
        {
            string directory=Path.Combine(Path.GetTempPath(),"ClearPlanAriaTest-"+Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var request=new AriaReportUploadRequest("SYNTHETIC_A","SYNTHETIC_A","Patient/test","plan-a","Organization/test",
                "urn:synthetic:type","review","Plan review",System.Text.Encoding.ASCII.GetBytes("%PDF-SYNTHETIC"),"synthetic.pdf",DateTimeOffset.UtcNow);
            try
            {
                using(var journal=new AriaUploadJournal(directory,"https://example.invalid/fhir/r4","SYNTHETIC_A","plan-a"))
                {
                    journal.Record("Pending",request);
                    bool locked=false;
                    try {using(var second=new AriaUploadJournal(directory,"https://example.invalid/fhir/r4","SYNTHETIC_A","plan-a")) {}}
                    catch(IOException){locked=true;}
                    TestAssert.True(locked,"Concurrent uploads must not acquire the same target-plan binding.");
                }
                bool blocked=false;
                try {using(var reopened=new AriaUploadJournal(directory,"https://example.invalid/fhir/r4","SYNTHETIC_A","plan-a")) {}}
                catch(InvalidOperationException){blocked=true;}
                TestAssert.True(blocked,"Restart must preserve an unresolved Pending receipt and prevent replay.");
                using(var other=new AriaUploadJournal(directory,"https://example.invalid/fhir/r4","SYNTHETIC_A","plan-b"))
                {
                    bool mismatched=false;
                    try {other.Record("Pending",request);} catch(ArgumentException){mismatched=true;}
                    TestAssert.True(mismatched,"A receipt must not record a different plan's request.");
                    var requestB=new AriaReportUploadRequest("SYNTHETIC_A","SYNTHETIC_A","Patient/test","plan-b","Organization/test",
                        "urn:synthetic:type","review","Plan review",request.PdfBytes,"synthetic.pdf",DateTimeOffset.UtcNow);
                    other.Record("Pending",requestB); other.Record("Verified",requestB,"DocumentReference/test");
                }
                using(var reopened=new AriaUploadJournal(directory,"https://example.invalid/fhir/r4","SYNTHETIC_A","plan-b"))
                {TestAssert.Equal("Verified",reopened.PreviousState);TestAssert.Equal("DocumentReference/test",reopened.PreviousResourceReference);}
                TestAssert.Equal(2,Directory.GetFiles(directory,"*.jsonl").Length);
                string sentDirectory=Path.Combine(directory,"sent-before-readback");
                using(var sent=new AriaUploadJournal(sentDirectory,"https://example.invalid/fhir/r4","SYNTHETIC_A","plan-a"))
                {sent.Record("Pending",request);sent.Record("Sent",request,"DocumentReference/test");}
                bool sentBlocked=false;
                try {using(var restart=new AriaUploadJournal(sentDirectory,"https://example.invalid/fhir/r4","SYNTHETIC_A","plan-a")) {}}
                catch(InvalidOperationException){sentBlocked=true;}
                TestAssert.True(sentBlocked,"A crash after receipt but before readback must remain blocked until reconciled.");
            }
            finally
            {
                if(Path.GetDirectoryName(directory)==Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar) &&
                    Path.GetFileName(directory).StartsWith("ClearPlanAriaTest-",StringComparison.Ordinal)) Directory.Delete(directory,true);
            }
        }

        public static void InMemoryReportForUpload()
        {
            var method=typeof(ClearPlan.Reporting.MigraDoc.ReportPdf).GetMethod("RenderToBytes");
            TestAssert.True(method!=null,"The upload must render before any remote filesystem write.");
            var document=new ClearPlan.Reporting.ReviewSnapshotReportMapper().Map(
                ClearPlan.Core.Simulation.SyntheticScenarioFactory.Create("baseline-pass"));
            var bytes=(byte[])method.Invoke(new ClearPlan.Reporting.MigraDoc.ReportPdf(),new object[]{document});
            TestAssert.True(bytes.Length>1000 && bytes.Length<AriaReportUploadRequest.MaximumPdfBytes);
            TestAssert.Equal("%PDF-",System.Text.Encoding.ASCII.GetString(bytes,0,5));
        }

        public static void CanceledFileWorkReleasesLateLease()
        {
            var entered=new System.Threading.Tasks.TaskCompletionSource<bool>();
            var release=new System.Threading.Tasks.TaskCompletionSource<bool>();
            var discarded=new System.Threading.Tasks.TaskCompletionSource<bool>();
            using(var cancel=new CancellationTokenSource())
            {
                var pending=AriaFileWork.RunAsync(()=> {entered.SetResult(true);release.Task.GetAwaiter().GetResult();return "owned-lease";},
                    cancel.Token,value=>{TestAssert.Equal("owned-lease",value);discarded.SetResult(true);});
                TestAssert.True(entered.Task.Wait(2000)); cancel.Cancel();
                bool canceled=false;
                try {pending.GetAwaiter().GetResult();}catch(OperationCanceledException){canceled=true;}
                TestAssert.True(canceled && !release.Task.IsCompleted,"UI cancellation must not wait for a blocked filesystem operation.");
                release.SetResult(true); TestAssert.True(discarded.Task.Wait(2000),"A late file lease must be disposed.");
            }
            TestAssert.Equal(7,AriaFileWork.RunAsync(()=>7,CancellationToken.None).GetAwaiter().GetResult());
        }

        private static JObject ValidJson()
        { return JObject.Parse("{\"SchemaVersion\":1,\"Enabled\":true,\"BaseUrl\":\"https://example.invalid/fhir/r4\",\"TokenUrl\":\"https://example.invalid/token\",\"ClientId\":\"synthetic-client\",\"DocumentTypeDisplay\":\"Plan review\"}"); }
        private static void Reject(Action action)
        {
            bool rejected=false;
            try {action();} catch(Exception exception)
            {
                rejected=true;
                TestAssert.True(exception.InnerException==null);
                TestAssert.False(exception.Message.Contains("SYNTHETIC_SECRET") || exception.Message.Contains("example.invalid"));
            }
            TestAssert.True(rejected,"Unsafe or unsupported ARIA configuration was accepted.");
        }
    }
}
