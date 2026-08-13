using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using VMS.TPS.Common.Model.API;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan
{
    public class PlanningItemDetailsCalculator
    {
        public ObservableCollection<PlanningItemDetailsViewModel> Calculate(PlanningItemViewModel activePlanningItem, ObservableCollection<PlanningItemViewModel> planningItemComboBoxList, 
            ObservableCollection<PQMSummaryViewModel> PqmSummaries, List<ErrorViewModel> ErrorGrid)
        {
            var PlanningItemSummaries = new ObservableCollection<PlanningItemDetailsViewModel>();
            var PlanSummary = new ObservableCollection<PlanningItemDetailsViewModel>();
            bool planIsChecked = false;
            bool isCCEnabled = false;
            foreach (PlanningItemViewModel planningItem in planningItemComboBoxList)
            {
                string pqmResult = "";
                string ccResult = "";
                string pcResult = "";
                string rpiResult = "";
                if (planningItem.PlanningItemIdWithCourse == activePlanningItem.PlanningItemIdWithCourse)
                {
                    int pqmsPassing = 0;
                    int pqmsTotal = 0;
                    if (PqmSummaries != null)
                    {
                        pqmsTotal = PqmSummaries.Count;
                        foreach (var row in PqmSummaries)
                        {
                            if (row.Met == "Goal" || row.Met == "Variation")
                                pqmsPassing += 1;
                        }
                    }
                    pqmResult = pqmsPassing.ToString() + "/" + pqmsTotal.ToString();
                    
                    
                    
                    int errorChecksPassing = 0;
                    int errorChecksTotal = ErrorGrid.Count();
                    foreach (var row in ErrorGrid)
                    {
                        if (row.Status == "3 - OK" || row.Status == "2 - Variation")
                            errorChecksPassing += 1;
                    }
                    pcResult = errorChecksPassing.ToString() + "/" + errorChecksTotal.ToString();
                }

                if (planningItem.PlanningItemObject is PlanSum)
                {
                    PlanSum planSum = (PlanSum)planningItem.PlanningItemObject;
                    int sumOfFractions = 0;
                    int sumOfTreatedFractions = 0;
                    string TreatmentApprover0="";
                    double sumOfPlanSetupDoses = 0;
                    int indexPlans = 0;
                    isCCEnabled = false;
                    if (planSum == activePlanningItem.PlanningItemObject)
                    {
                        planIsChecked = true;
                    }
                    else
                    {
                        planIsChecked = false;
                    }

                    foreach (PlanSetup planSetup in planSum.PlanSetups.OrderBy(x => x.CreationDateTime))
                    {
                        indexPlans++;
                        string planTarget;
                        sumOfFractions += planSetup.NumberOfFractions.Value;
                        sumOfTreatedFractions += planSetup.TreatmentSessions.Where(x => x.Status.ToString().ToLower() == "completed").Count();
                        TreatmentApprover0 += indexPlans + ". " + (planSetup.TreatmentApproverDisplayName==""?"not approved": planSetup.TreatmentApproverDisplayName+" ("+planSetup.Id.Substring(0, planSetup.Id.IndexOf(" ")!=-1? planSetup.Id.IndexOf(" "): planSetup.Id.Length) +")")  + "  ";
                        sumOfPlanSetupDoses += planSetup.TotalDose.Dose;
                        if (planSetup.TargetVolumeID != null)
                            planTarget = planSetup.TargetVolumeID;
                        else
                            planTarget = "No target selected";
                    }
                    var PlanSumSummary = new PlanningItemDetailsViewModel
                    {
                        PQM = planIsChecked,
                        CC = isCCEnabled,
                        PlanningItemIdWithCourse = planSum.Course + "/" + planSum.Id,
                        PlanningItemCourseId = planSum.Course.Id,
                        ApprovalStatus = "PlanSum",
                        PlanningItemObject = planningItem.PlanningItemObject,
                        PlanName = planSum.Course + "/" + planSum.Id,
                        //PlanCreated = planSum.HistoryDateTime.ToString(),
                        PlanCreated = DateTime.Parse(planSum.HistoryDateTime.ToString()).ToString("yyyy-MM-dd HH:mm:ss"),
                        PlanFractions = sumOfFractions.ToString(),
                        TreatedSessions = sumOfTreatedFractions.ToString(),
                        PlanTotalDose = sumOfPlanSetupDoses.ToString("0.##"),
                        PQMResult = pqmResult,
                        CCResult = ccResult,
                        PCResult = pcResult,
                        RPIResult = rpiResult,
                        TreatmentApprover = TreatmentApprover0,
                        PlanCount = indexPlans
                    };
                    PlanningItemSummaries.Add(PlanSumSummary);

                }
                else  //planningitem is plansetup
                {
                    PlanSetup planSetup = (PlanSetup)planningItem.PlanningItemObject;
                    if (planSetup == activePlanningItem.PlanningItemObject)
                    {
                        planIsChecked = true;
                        isCCEnabled = true;
                    }
                    else
                    {
                        planIsChecked = false;
                        isCCEnabled = false;
                    }
                    var approvalStatus = "";
                    if (planSetup.PlanIntent == "VERIFICATION")
                    {
                        approvalStatus = "VerificationPlan";
                    }
                    else
                        approvalStatus = planSetup.ApprovalStatus.ToString();

                    bool extended = false;
                    bool bolus = false;
                    foreach (Beam b in planSetup.Beams)
                    {
                        if (b.IsGantryExtended)
                            extended = true;
                        if(b.Boluses.Count() > 0)
                            bolus = true;
                    }
                    string planTarget;
                    if (planSetup.TargetVolumeID != null)
                        planTarget = planSetup.TargetVolumeID;
                    else
                        planTarget = "No target selected";
                    var PlanningItemSummary = new PlanningItemDetailsViewModel
                    {
                        PQM = planIsChecked,
                        CC = isCCEnabled,
                        PlanningItemIdWithCourse = planSetup.Course + "/" + planSetup.Id,
                        PlanningItemCourseId = planSetup.Course.Id,
                        ApprovalStatus = approvalStatus,
                        UseGated = planSetup.UseGating.ToString().ToLower(),
                        UseExtended = extended.ToString().ToLower(),
                        UseBolus = bolus.ToString().ToLower(),
                        PlanName = planSetup.Course + "/" + planSetup.Id,
                        PlanningItemObject = planningItem.PlanningItemObject,
                        //PlanCreated = planSetup.CreationDateTime.ToString(),
                        PlanCreated = DateTime.Parse(planSetup.CreationDateTime.ToString()).ToString("yyyy-MM-dd HH:mm:ss"),
                        PlanFxDose = planSetup.DosePerFraction.Dose.ToString("0.##"),
                        PlanFractions = planSetup.NumberOfFractions.ToString(),
                        TreatedSessions = planSetup.TreatmentSessions.Where(x => x.Status.ToString().ToLower() == "completed").Count().ToString(),
                        PlanTotalDose = planSetup.TotalDose.Dose.ToString("0.##"),
                        PlanTarget = planTarget,
                        PQMResult = pqmResult,
                        CCResult = ccResult,
                        PCResult = pcResult,
                        RPIResult = rpiResult,
                    };
                    PlanningItemSummaries.Add(PlanningItemSummary);
                }
            }
            return PlanningItemSummaries;
        }
    }
}
