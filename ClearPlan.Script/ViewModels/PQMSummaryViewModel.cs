using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Media;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan
{
    public class PQMSummaryViewModel : ViewModelBase
    {
        public string ConstraintId { get; set; }
        public string TemplateId { get; set; }
        public string[] TemplateCodes { get; set; }
        public string[] TemplateAliases { get; set; }
        public string[] TemplateType { get; set; }
        public ObservableCollection<StructureViewModel> StructureList { get; set; }
        //public ObservableCollection<PlanningItemViewModel> planningItemComboBoxList { get; set; }
        public string StructVolume { get; set; }
        public string StructType { get; set; }
        public string DVHObjective { get; set; }
        public string Goal { get; set; }
        public string Achieved { get; set; }
        public string AchievedComparison { get; set; }
        public string Met { get; set; }
        private bool _accept;
        public bool Accept
        {
            get { return _accept; }
            set
            {
                if (_accept == value)
                {
                    return;
                }

                _accept = value;
                NotifyPropertyChanged("Accept");
            }
        }

        private bool _ignore;
        public bool Ignore
        {
            get { return _ignore; }
            set
            {
                if (_ignore == value)
                {
                    return;
                }

                _ignore = value;
                NotifyPropertyChanged("Ignore");
            }
        }
        public string Variation { get; set; }
        public string Priority { get; set; }
        public string Source { get; set; }
        public string Comment { get; set; }
        public bool isCalculated { get; set; }
        public double AchievedPercentageOfGoal { get; set; }
        public string StructureName { get; set; }
        public string StructureNameWithCode { get; set; }
        public string MappedStructureDisplay
        {
            get
            {
                if (_Structure != null &&
                    !string.IsNullOrWhiteSpace(
                        _Structure.StructureNameWithCode))
                {
                    return _Structure.StructureNameWithCode;
                }

                if (!string.IsNullOrWhiteSpace(StructureNameWithCode))
                {
                    return StructureNameWithCode;
                }

                return !string.IsNullOrWhiteSpace(StructureName)
                    ? StructureName
                    : "nicht zugeordnet";
            }
        }
        public SolidColorBrush AchievedColor { get; set; }       
        public PlanningItemViewModel ActivePlanningItem { get; set;}

        public StructureViewModel _Structure;
        public StructureViewModel Structure
        {
            get { return _Structure; }
            set
            {
                ApplyStructureEvaluation(EvaluateStructure(value));
            }
        }

        internal PqmStructureEvaluation EvaluateStructure(
            StructureViewModel structure)
        {
            var evaluation = new PqmStructureEvaluation
            {
                Structure = structure
            };
            if (structure == null || Goal == null)
            {
                return evaluation;
            }

            var calculator = new PQMSummaryCalculator();
            string achieved = calculator.CalculateMetric(
                ActivePlanningItem.PlanningItemStructureSet,
                structure,
                ActivePlanningItem,
                DVHObjective,
                Variation);
            string met = calculator.EvaluateMetric(
                achieved,
                Goal,
                Variation);
            var color = Calculators.PQMColors.GetAchievedColor(
                structure.Structure,
                Goal,
                DVHObjective,
                achieved);

            evaluation.HasCalculatedMetrics = true;
            evaluation.StructVolume = structure.VolumeValue;
            evaluation.Achieved = achieved;
            evaluation.Met = met;
            evaluation.StructureName = structure.StructureName;
            evaluation.StructureNameWithCode =
                structure.StructureNameWithCode;
            evaluation.AchievedColor = color.Item1;
            evaluation.AchievedPercentageOfGoal = color.Item2;
            return evaluation;
        }

        internal void ApplyStructureEvaluation(
            PqmStructureEvaluation evaluation)
        {
            if (evaluation == null)
            {
                throw new ArgumentNullException("evaluation");
            }

            _Structure = evaluation.Structure;
            NotifyPropertyChanged("Structure");
            NotifyPropertyChanged("MappedStructureDisplay");
            if (!evaluation.HasCalculatedMetrics)
            {
                return;
            }

            StructVolume = evaluation.StructVolume;
            Achieved = evaluation.Achieved;
            Met = evaluation.Met;
            StructureName = evaluation.StructureName;
            StructureNameWithCode =
                evaluation.StructureNameWithCode;
            AchievedColor = evaluation.AchievedColor;
            AchievedPercentageOfGoal =
                evaluation.AchievedPercentageOfGoal;
            NotifyPropertyChanged("StructVolume");
            NotifyPropertyChanged("Achieved");
            NotifyPropertyChanged("Met");
            NotifyPropertyChanged("StructureName");
            NotifyPropertyChanged("StructureNameWithCode");
            NotifyPropertyChanged("MappedStructureDisplay");
            NotifyPropertyChanged("AchievedColor");
            NotifyPropertyChanged("AchievedPercentageOfGoal");
        }

        internal PqmStructureEvaluation CaptureStructureEvaluation()
        {
            return new PqmStructureEvaluation
            {
                Structure = _Structure,
                HasCalculatedMetrics = true,
                StructVolume = StructVolume,
                Achieved = Achieved,
                Met = Met,
                StructureName = StructureName,
                StructureNameWithCode = StructureNameWithCode,
                AchievedColor = AchievedColor,
                AchievedPercentageOfGoal =
                    AchievedPercentageOfGoal
            };
        }
    }

    internal sealed class PqmStructureEvaluation
    {
        public StructureViewModel Structure { get; set; }
        public bool HasCalculatedMetrics { get; set; }
        public string StructVolume { get; set; }
        public string Achieved { get; set; }
        public string Met { get; set; }
        public string StructureName { get; set; }
        public string StructureNameWithCode { get; set; }
        public SolidColorBrush AchievedColor { get; set; }
        public double AchievedPercentageOfGoal { get; set; }
    }
}
