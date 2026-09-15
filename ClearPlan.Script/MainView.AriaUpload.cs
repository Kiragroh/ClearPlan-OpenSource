using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using ClearPlan.Core.Integration;
using ClearPlan.Core.Review;
using ClearPlan.Presentation.ViewModels;
using ClearPlan.Reporting;
using ClearPlan.Reporting.MigraDoc;

namespace ClearPlan
{
    public partial class MainView
    {
        private CancellationTokenSource ariaPreparation;
        internal void CancelAriaPreparation()
        { if(ariaPreparation!=null) ariaPreparation.Cancel(); }

        internal void RefreshAriaAvailability(ReviewWorkspaceViewModel workspace)
        {
            if(workspace==null) return;
            bool privacy=Anonymize_CheckBox.IsChecked==true;
            bool configured=!string.IsNullOrWhiteSpace(_settings.Paths.AriaUploadConfigJsonPath);
            workspace.SetAriaUploadAvailability(configured && !privacy,ariaPreparation!=null,
                privacy ? "Privacy-Modus: kein Versand in die Patientenakte. Für den Versand die klinische Identität anzeigen." :
                configured ? "Aktuellen PDF-Bericht vorbereiten. Patient, Plan und Dokumenttyp werden vor dem Senden bestätigt." :
                "Unter Einstellungen den Pfad zur lokalen ARIA-Verbindungskonfiguration hinterlegen.");
        }

        private bool AriaContextIsCurrent(ReviewSnapshot snapshot,ReviewWorkspaceViewModel workspace,string patientId,
            string interactiveUserId,string activePlanName)
        {
            return _clinicalReviewHost!=null && _clinicalReviewHost.CanExportCurrentSnapshot(snapshot) &&
                ReferenceEquals(workspace,_clinicalReviewHost.CurrentViewModel) && !snapshot.Synthetic &&
                Anonymize_CheckBox.IsChecked!=true && _vm.Patient!=null && string.Equals(_vm.Patient.Id,patientId,StringComparison.Ordinal) &&
                _vm.User!=null && string.Equals(_vm.User.Id,interactiveUserId,StringComparison.Ordinal) &&
                _vm.ActivePlanningItem!=null && string.Equals(_vm.ActivePlanningItem.PlanningItemId,activePlanName,StringComparison.Ordinal);
        }

        internal async Task HandleSharedAriaUploadAsync(ReviewSnapshot snapshot,ReviewWorkspaceViewModel workspace)
        {
            if(ariaPreparation!=null || sharedReportExportRunning || snapshot==null || workspace==null || !workspace.CanUploadToAria) return;
            string patientId=_vm.Patient==null ? null : _vm.Patient.Id;
            // _vm.User is the ScriptContext.CurrentUser captured by the native entry point, not the OAuth client.
            string interactiveUserId=_vm.User==null ? null : _vm.User.Id;
            string activePlanName=_vm.ActivePlanningItem==null ? null : _vm.ActivePlanningItem.PlanningItemId;
            if(string.IsNullOrWhiteSpace(interactiveUserId))
            {ShowAriaStatus("ARIA-Autor fehlt: ClearPlan mit dem persönlichen ARIA/ESAPI-Benutzer neu öffnen. Kein Bericht gesendet.",AnalysisStatusSeverity.Warning);return;}
            if(!AriaContextIsCurrent(snapshot,workspace,patientId,interactiveUserId,activePlanName)) return;
            // Copy patient strings while on the ESAPI thread; all subsequent HTTP objects are detached.
            string patientLabel=_vm.Patient.LastName+", "+_vm.Patient.FirstName+" | ID: "+patientId;
            string configPath=_settings.ResolvePath(_settings.Paths.AriaUploadConfigJsonPath);
            string outputRoot=Path.Combine(_settings.ResolvePath(_settings.Paths.ReportsDirectory),"AriaUploads");
            ariaPreparation=new CancellationTokenSource(TimeSpan.FromMinutes(6));
            var cancellation=ariaPreparation.Token;
            RefreshAriaAvailability(workspace);
            bool attempted=false;
            try
            {
                ShowAriaStatus("ARIA: Verbindung und Dokumenttyp werden geprüft …",AnalysisStatusSeverity.Info);
                var config=await AriaUploadConfiguration.LoadAsync(configPath,cancellation);
                if(!config.Enabled) {ShowAriaStatus("ARIA-Versand ist in der lokalen Konfiguration deaktiviert.",AnalysisStatusSeverity.Warning); return;}
                using(var tokenHttp=AriaCertificatePolicy.CreateClient(new Uri(config.TokenUrl),config.TokenCertificateSha256,config.TimeoutSeconds))
                using(var fhirHttp=AriaCertificatePolicy.CreateClient(new Uri(config.BaseUrl),config.FhirCertificateSha256,config.TimeoutSeconds))
                {
                    var auth=new AriaOAuthTokenProvider(tokenHttp,new Uri(config.TokenUrl),config.ClientId,config.Scope,
                        ct=>config.LoadSecretAsync(configPath,ct));
                    var attachmentReader=new AriaAttachmentReader(config.AttachmentReadbackRoots);
                    var client=new AriaFhirClient(fhirHttp,new Uri(config.BaseUrl),auth.GetTokenAsync,config.TimeoutSeconds,
                        config.AttachmentReadbackRoots.Length==0 ? null : new Func<string,CancellationToken,Task<byte[]>>(attachmentReader.ReadAsync));
                    var author=await client.ResolveAuthorAsync(interactiveUserId,cancellation);
                    var provider=await client.ResolveProviderAsync(config.ProviderReference,cancellation);
                    var type=await client.ResolveDocumentTypeAsync(provider.Reference,config.DocumentTypeCode,config.DocumentTypeDisplay,cancellation);
                    if(!AriaContextIsCurrent(snapshot,workspace,patientId,interactiveUserId,activePlanName)) throw new OperationCanceledException();
                    var patient=await client.ResolvePatientAsync(patientId,config.PatientIdentifierSystem,cancellation);
                    if(!AriaContextIsCurrent(snapshot,workspace,patientId,interactiveUserId,activePlanName)) throw new OperationCanceledException();
                    ShowAriaStatus("ARIA: aktueller Planreport wird vorbereitet …",AnalysisStatusSeverity.Info);
                    var analysis=workspace.IncludeBeamEyeViews ? await _clinicalReviewHost.PrepareReportBevsAsync(snapshot) : snapshot.PlanAnalysis;
                    await _clinicalReviewHost.PreparePlanImagesAsync(snapshot);
                    cancellation.ThrowIfCancellationRequested();
                    if(!AriaContextIsCurrent(snapshot,workspace,patientId,interactiveUserId,activePlanName)) throw new OperationCanceledException();
                    foreach(var series in snapshot.DvhSeries)
                    {
                        var selection=workspace.DvhSeries.FirstOrDefault(row=>row.StableId==series.StableId);
                        if(selection!=null) series.Selected=selection.IsSelected;
                    }
                    ClearPlan.Presentation.Views.CollisionView.CaptureReportSweep(workspace.Collision);
                    var document=new ReviewSnapshotReportMapper().Map(snapshot);
                    document.PlanAnalysis=analysis; document.PatientDisplayLabel=patientLabel;
                    ApplyReportOptions(document,workspace);
                    string title="ClearPlan-plan-review-"+DateTime.UtcNow.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N").Substring(0,8)+".pdf";
                    string pdfPath=Path.Combine(outputRoot,title);
                    await System.Windows.Threading.Dispatcher.Yield(System.Windows.Threading.DispatcherPriority.Background);
                    byte[] pdf=new ReportPdf().RenderToBytes(document);
                    // Native image rendering stays on the owner STA; network filesystem I/O does not.
                    await AriaFileWork.RunAsync(()=> {Directory.CreateDirectory(outputRoot); File.WriteAllBytes(pdfPath,pdf);return true;},cancellation);
                    // Keep the report's printed provenance instant. Only the separately disclosed ARIA date may differ.
                    var createdUtc=document.GeneratedUtc.ToUniversalTime();
                    var documentDateUtc=AriaDocumentDatePolicy.Calculate(createdUtc,DateTimeOffset.UtcNow,config.DocumentDateSafetyMinutes);
                    var request=new AriaReportUploadRequest(patientId,patient.Identifier,patient.Reference,snapshot.ActivePlanKey,
                        provider.Reference,type.System,type.Code,type.Display,pdf,title,createdUtc,
                        "ClearPlan plan review - preliminary",config.CategoryCode,config.CategoryCode,documentDateUtc)
                        .WithClinicalMetadata(activePlanName,author.Reference,interactiveUserId);
                    var journal=await AriaFileWork.RunAsync(()=>new AriaUploadJournal(Path.Combine(outputRoot,"Receipts"),config.BaseUrl,patientId,snapshot.ActivePlanKey),cancellation,j=>j.Dispose());
                    try
                    {
                        cancellation.ThrowIfCancellationRequested();
                        if(!AriaContextIsCurrent(snapshot,workspace,patientId,interactiveUserId,activePlanName)) throw new OperationCanceledException();
                        string previous=journal.PreviousState=="Verified"
                            ? "\nFür diesen Plan wurde bereits ein Bericht gesendet. Hiermit wird bewusst ein neuer Bericht angelegt.\n" : "";
                        string dates=request.DocumentDateUtc != request.CreatedUtc
                            ? "\nARIA-Dokumentdatum (UTC): "+request.DocumentDateUtc.ToString("yyyy-MM-dd HH:mm:ss 'Z'",CultureInfo.InvariantCulture)+
                                "\nBericht erstellt (UTC): "+request.CreatedUtc.ToString("yyyy-MM-dd HH:mm:ss 'Z'",CultureInfo.InvariantCulture)+
                                "\nKonfigurierte ARIA-Zeitreserve: "+config.DocumentDateSafetyMinutes+" min. PDF-Inhalt und Erstellzeit bleiben unverändert."
                            : "";
                        string confirmation="Diesen PDF-Bericht als vorläufiges Dokument in ARIA ablegen?\n\n"+
                            patientLabel+"\nPlan: "+snapshot.PlanDisplayLabel+"\nARIA: "+patient.Reference+
                            "\nEinrichtung: "+provider.Display+" ("+provider.Reference+")\nDokumenttyp: "+type.Display+
                            "\nTemplateName: "+request.TemplateName+"\nAutor: "+interactiveUserId+" ("+author.Reference+")"+
                            "\nPDF: "+title+"\nSHA-256: "+request.PdfSha256+dates+previous+
                            "\n\nDiese Aktion schreibt ein Dokument in die Patientenakte. Bestrahlungsplan und Dosis bleiben unverändert.";
                        if(MessageBox.Show(Window.GetWindow(this),confirmation,"ClearPlan · ARIA-Versand bestätigen",MessageBoxButton.YesNo,
                            MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)
                        {ShowAriaStatus("ARIA-Versand abgebrochen. Vorbereitete PDF bleibt im Reportordner.",AnalysisStatusSeverity.Info);return;}
                        cancellation.ThrowIfCancellationRequested();
                        if(!AriaContextIsCurrent(snapshot,workspace,patientId,interactiveUserId,activePlanName)) throw new OperationCanceledException();
                        // Persist intent before POST. A crash or uncertain response remains blocked across GUI/Runner restarts.
                        await AriaFileWork.RunAsync(()=>{journal.Record("Pending",request);return true;},CancellationToken.None);
                        if(cancellation.IsCancellationRequested || !AriaContextIsCurrent(snapshot,workspace,patientId,interactiveUserId,activePlanName))
                        {
                            // This branch has not entered UploadAsync: a known local cancellation is not an uncertain server write.
                            await AriaFileWork.RunAsync(()=>{journal.Record("Rejected",request);return true;},CancellationToken.None);
                            throw new OperationCanceledException();
                        }
                        attempted=true;
                        ShowAriaStatus("ARIA: Bericht wird einmalig gesendet …",AnalysisStatusSeverity.Info);
                        var result=await client.UploadAsync(request,cancellation);
                        if(result.State==AriaUploadState.Rejected)
                        {await AriaFileWork.RunAsync(()=>{journal.Record("Rejected",request);return true;},CancellationToken.None);ShowAriaStatus(result.Message,AnalysisStatusSeverity.Error);return;}
                        await AriaFileWork.RunAsync(()=>{journal.Record(result.State==AriaUploadState.Sent ? "Sent" : "Uncertain",request,result.ResourceReference);return true;},CancellationToken.None);
                        if(string.IsNullOrEmpty(result.ResourceReference))
                        {ShowAriaStatus("ARIA-Ergebnis unklar ("+result.ProcessingStageCode+"). Nicht erneut senden; Patientenakte und gespeicherten Beleg prüfen.",AnalysisStatusSeverity.Warning);return;}
                        // A validated Location survives an incomplete create response. Recover with GET only, never a second POST.
                        ShowAriaStatus("ARIA: angelegte Dokumentreferenz und gespeicherte PDF werden zurückgelesen …",AnalysisStatusSeverity.Info);
                        var verification=await client.VerifyAsync(result.ResourceReference,request,cancellation);
                        if(verification.State==AriaVerificationState.Verified)
                        {
                            await AriaFileWork.RunAsync(()=>{journal.Record("Verified",request,result.ResourceReference);return true;},CancellationToken.None);
                            ShowAriaStatus("ARIA: Patient, TemplateName, Autor, Dokumentdatum und PDF-Fingerprint bestätigt."+
                                (verification.AttachmentCreationChecked ? " Erstellzeit bestätigt." : " Erstellzeit von ARIA nicht geliefert und nicht geprüft."),AnalysisStatusSeverity.Success);
                        }
                        else
                        {await AriaFileWork.RunAsync(()=>{journal.Record("Uncertain",request,result.ResourceReference);return true;},CancellationToken.None);ShowAriaStatus("ARIA-Dokument nicht vollständig rückgeprüft. Nicht erneut senden; gespeicherten Beleg und Patientenakte prüfen.",AnalysisStatusSeverity.Warning);}
                    }
                    finally{AriaFileWork.DisposeLater(journal);}
                }
            }
            catch(OperationCanceledException)
            {ShowAriaStatus(attempted ? "ARIA-Ergebnis nach Abbruch unklar. Gespeicherten Beleg und Patientenakte prüfen; nicht erneut senden." : "ARIA-Vorbereitung abgebrochen oder Plankontext geändert. Kein Bericht gesendet.",AnalysisStatusSeverity.Warning);}
            catch(AriaAuthorResolutionException)
            {ShowAriaStatus("ARIA-Autor nicht eindeutig auflösbar. ARIA-Administration: persönlichen ESAPI-Benutzer mit aktivem Practitioner/Staff verknüpfen und system/Practitioner.rs freigeben. Kein Bericht gesendet; kein Ersatz durch das Dienstkonto.",AnalysisStatusSeverity.Warning);}
            catch(Exception)
            {ShowAriaStatus(attempted ? "ARIA-Ergebnis nicht vollständig bestätigt. Beleg im Reportordner prüfen; nicht erneut senden." : "ARIA-Vorbereitung fehlgeschlagen oder ein früherer Versand ist ungeklärt. Verbindung, Zuordnung und Belege im Reportordner prüfen. Kein neuer Bericht gesendet.",AnalysisStatusSeverity.Warning);}
            finally
            {
                ariaPreparation.Dispose(); ariaPreparation=null;
                if(_clinicalReviewHost!=null) RefreshAriaAvailability(_clinicalReviewHost.CurrentViewModel);
            }
        }

        private void ShowAriaStatus(string message,AnalysisStatusSeverity severity)
        {
            SharedWorkspaceStatusText.Text=message; SharedWorkspaceStatusText.Tag=severity.ToString();
            SharedWorkspaceStatusText.ToolTip="ARIA-Dokumentversand ist eine gesonderte Schreibaktion in der Patientenakte. Kein Ändern von Plan, Dosis oder Strukturen über ESAPI.";
        }
    }
}
