using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using VMS.TPS.Common.Model.API;

namespace ClearPlan
{
    public class PlanningItemViewModel : ViewModelBase
    {
        public string PlanningItemId { get; set; }
        public string PlanningItemCourse { get; set; }
        public string PlanningItemIdWithCourse { get; set; }
        //public string PlanningItemCourseId { get; set; }
        public string PlanningItemTargetId { get; set; }
        public string PlanningItemFx { get; set; }
        public string PlanningItemTreatmentApproverDisplayName { get; set; }
        public string PlanningItemIdWithCourseAndType { get; set; }
        public string PlanningItemType { get; set; }
        public string Algorithmus { get; set; }
        public string Technique { get; set; }
        public string Orientation { get; set; }
        public string PlanningGatingExtenedString  { get; set; }
        public string PlanningItemUID { get; set; }
        public int PlanningItemSumApproved { get; set; }
        public string PlanningApprover { get; set; }
        public double PlanningItemTreatmentPercentage { get; set; }
        public double PlanningItemDoseFx { get; set; }
        public double PlanningItemDose { get; set; }
        
        public PlanningItem PlanningItemObject { get; set; }
        public IEnumerable<PlanSetup> PlanningItemPlans { get; set; }
        public List<Beam> PlanningItemBeams { get; set; }
        public PlanSetup PlanningItemPlanSetup { get; set; }
        public DateTime Creation { get; set; }
        public string PlanningItemTreatmentApproveDate { get; set; }
        public string PlanningItemApproveDate { get; set; }
        public StructureSet PlanningItemStructureSet { get; set; }
        public Image PlanningItemImage { get; set; }
        
        public PlanningItemViewModel(PlanningItem planningItem)
        {
            if (planningItem is PlanSetup)
            {
                PlanSetup planSetup = (PlanSetup)planningItem;
                PlanningItemPlanSetup = planSetup;
                PlanningItemId = planSetup.Id;
                PlanningItemCourse = planSetup.Course.Id;
                PlanningItemIdWithCourse = PlanningItemCourse + "/" + PlanningItemId;
                //PlanningItemCourseId = PlanningItemCourse;
                PlanningItemType = "Plan";
                PlanningItemTreatmentPercentage = planSetup.TreatmentPercentage;
                PlanningItemUID = planSetup.UID;
                PlanningItemTreatmentApproverDisplayName = planSetup.TreatmentApproverDisplayName;
                //var creation = (DateTime)patientSummary.CreationDateTime;
                //var creationdate = creation.ToString("yyyy-MM-dd");
                PlanningItemTreatmentApproveDate = planSetup.TreatmentApprovalDate!= "" ? DateTime.Parse(planSetup.TreatmentApprovalDate).ToString("yyyy-MM-dd HH:mm"): planSetup.TreatmentApprovalDate;
                PlanningItemApproveDate = planSetup.PlanningApprovalDate != "" ? DateTime.Parse(planSetup.PlanningApprovalDate).ToString("yyyy-MM-dd HH:mm") : planSetup.PlanningApprovalDate;

                PlanningItemBeams = planSetup.Beams.FirstOrDefault() != null? planSetup.Beams.OrderBy(b => b.BeamNumber).ToList(): null;
                PlanningItemPlans = Enumerable.Empty<PlanSetup>();
                Orientation = " | Orientation: " +planSetup.StructureSet.Image.ImagingOrientation;
                PlanningItemIdWithCourseAndType = PlanningItemCourse + "/" + PlanningItemId + " (" + PlanningItemType + ")";
                PlanningItemObject = planSetup;
                PlanningItemDoseFx = planSetup.DosePerFraction.Dose;
                PlanningItemDose = planSetup.TotalDose.Dose;
                PlanningItemStructureSet = planSetup.StructureSet;
                PlanningItemTargetId = planSetup.TargetVolumeID;
                PlanningItemImage = planSetup.StructureSet.Image;
                PlanningItemSumApproved = 0;
                PlanningItemFx = planSetup.NumberOfFractions.ToString();
                Creation = (DateTime) planSetup.HistoryDateTime;
                PlanningGatingExtenedString = (planSetup.UseGating == true ? "Gating: ON, " : "Gating: OFF, ") 
                    + (planSetup.Beams.Where(x => x.IsGantryExtended).Any() == true ? "GantryExtended: Yes, " : "GantryExtended: No, ")
                    + (planSetup.Beams.Where(x => x.Boluses.Count()>0).Any() == true ? "Bolus: Used" : "Bolus: No");
                
                    Algorithmus = planSetup.PhotonCalculationModel.ToString() +  (planSetup.PhotonCalculationOptions.Where(x => x.Key == "CalculationGridSizeInCM").Count()>0 ? " (" + planSetup.PhotonCalculationOptions.Where(x => x.Key == "CalculationGridSizeInCM").FirstOrDefault().Value.ToString() + "cm)":"");
                try
                {
                    Technique = planSetup.Beams.Where(x => !x.IsSetupField).FirstOrDefault() == null ? "" : (planSetup.Beams.Where(x => !x.IsSetupField).Any(b => b.MLCPlanType.ToString().ToLower() == "static") && planSetup.Beams.Where(x => !x.IsSetupField).Any(b => b.MLCPlanType.ToString().ToLower() == "vmat")) ?
                        "Hybrid" : (planSetup.Beams.Where(x => !x.IsSetupField).Any(b => b.MLCPlanType.ToString().ToLower() == "static") ? "Static"
                        : (planSetup.Beams.Where(x => !x.IsSetupField).Any(b => b.MLCPlanType.ToString().ToLower() == "vmat") ? "VMAT"
                        : planSetup.Beams.Where(x => !x.IsSetupField).FirstOrDefault().MLCPlanType.ToString()));
                }
                catch {
                    //Algorithmus = "test";
                    Technique = "Error";
                }
                PlanningApprover = planSetup.PlanningApproverDisplayName==""?"":" | Planner: " + planSetup.PlanningApproverDisplayName;
    }
            if (planningItem is PlanSum)
            {
                PlanSum planSum = (PlanSum)planningItem;
                int countApprovals = 0;

                foreach(PlanSetup ps in planSum.PlanSetups)
                {
                    if (ps.ApprovalStatus != VMS.TPS.Common.Model.Types.PlanSetupApprovalStatus.UnApproved &&
                        ps.ApprovalStatus != VMS.TPS.Common.Model.Types.PlanSetupApprovalStatus.Rejected &&
                        ps.ApprovalStatus != VMS.TPS.Common.Model.Types.PlanSetupApprovalStatus.Reviewed)
                        countApprovals++;
                }

                PlanningItemId = planSum.Id;
                PlanningItemCourse = planSum.Course.Id;
                PlanningItemIdWithCourse = PlanningItemCourse + "/" + PlanningItemId;
                PlanningItemType = "PlanSum";
                PlanningItemTreatmentApproverDisplayName = " ";
                PlanningItemIdWithCourseAndType = PlanningItemCourse + "/" + PlanningItemId + " (" + PlanningItemType + ")";
                PlanningItemObject = planSum;
                PlanningItemUID = "";
                //PlanningItemDoseFx = "PlanSum";
                //PlanningItemDose = "PlanSum";
                PlanningItemPlans = planSum.PlanSetups;
                PlanningItemSumApproved = countApprovals;
                PlanningItemTargetId = "PTV";
                PlanningItemFx = "PlanSum";
                PlanningItemStructureSet = planSum.StructureSet;
                PlanningItemImage = planSum.StructureSet.Image;
                Creation = (DateTime)planSum.HistoryDateTime;
                Orientation = " | Orientation: " + planSum.StructureSet.Image.ImagingOrientation;
            }
        }
    }
}
