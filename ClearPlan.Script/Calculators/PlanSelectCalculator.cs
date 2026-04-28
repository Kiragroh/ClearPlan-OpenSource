using System.Collections.ObjectModel;
using System.Linq;
using VMS.TPS.Common.Model.API;
using System;

namespace ClearPlan
{
    public class PlanSelectCalculator
    {
        public ObservableCollection<PlanSelectDetailViewModel> Calculate(ObservableCollection<PlanningItemViewModel> planningItemList)
        {
            var PlanningItemSummaries = new ObservableCollection<PlanSelectDetailViewModel>();
            var PlanSummary = new ObservableCollection<PlanSelectDetailViewModel>();
            foreach (PlanningItemViewModel planningItem in planningItemList)
            {
                if (planningItem.PlanningItemObject is PlanSum)
                {
                    PlanSum planSum = (PlanSum)planningItem.PlanningItemObject;
                    int sumOfFractions = 0;
                    double sumOfPlanSetupDoses = 0;
                    
                    foreach (PlanSetup planSetup in planSum.PlanSetups.OrderBy(x => x.CreationDateTime))
                    {
                        sumOfFractions += planSetup.NumberOfFractions.Value;
                        sumOfPlanSetupDoses += planSetup.TotalDose.Dose;
                    }
                    var PlanningItemSummary = new PlanSelectDetailViewModel
                    {
                        ActivePlanningItem = new PlanningItemViewModel(planSum),
                        PlanningItemIdWithCourse = planSum.Course + "/" + planSum.Id,
                        PlanningItemCourseId = planSum.Course.Id,
                        ApprovalStatus = "PlanSum",
                        PlanName = planSum.Course + "/" + planSum.Id,
                        PlanStructureSet = planSum.StructureSet.Id,
                        IsPlanSum = true,
                        PlanCreated = DateTime.Parse(planSum.HistoryDateTime.ToString()).ToString("yyyy-MM-dd HH:mm:ss"),

                        PlanFractions = sumOfFractions.ToString(),

                        PlanTotalDose = sumOfPlanSetupDoses.ToString("0.##"),

                        IsDoseValid = planSum.IsDoseValid()
                    };
                    PlanningItemSummaries.Add(PlanningItemSummary);
                }
                else
                {
                    PlanSetup planSetup = (PlanSetup)planningItem.PlanningItemObject;                       
                    string planTarget;
                    if (planSetup.TargetVolumeID != null)
                        planTarget = planSetup.TargetVolumeID;
                    else
                        planTarget = "No target selected";
                    var PlanningItemSummary = new PlanSelectDetailViewModel
                    {
                        ActivePlanningItem = new PlanningItemViewModel(planSetup),
                        PlanningItemIdWithCourse = planSetup.Course + "/" + planSetup.Id,
                        PlanningItemCourseId = planSetup.Course.Id,
                        ApprovalStatus = planSetup.ApprovalStatus.ToString(),
                        
                        PlanName = planSetup.Course + "/" + planSetup.Id,
                        IsPlanSum = false,
                        PlanStructureSet = planSetup.StructureSet.Id,
                        PlanNumFields = planSetup.Beams.Count().ToString(),
                        PlanningItemObject = planningItem.PlanningItemObject,
                        PlanCreated = DateTime.Parse(planSetup.CreationDateTime.ToString()).ToString("yyyy-MM-dd HH:mm:ss"),

                        PlanFxDose = planSetup.DosePerFraction.Dose.ToString("0.##"),

                        PlanFractions = planSetup.NumberOfFractions.ToString(),

                        PlanTotalDose = planSetup.TotalDose.Dose.ToString("0.##"),
                        PlanTarget = planTarget,
                        IsDoseValid = planSetup.IsDoseValid()
                    };
                    PlanningItemSummaries.Add(PlanningItemSummary);
                }
            }
            return PlanningItemSummaries;
        }
    }
}
