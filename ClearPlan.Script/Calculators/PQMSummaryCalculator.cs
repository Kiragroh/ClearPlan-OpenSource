using ClearPlan.Calculators;
using ClearPlan.Core.Constraints;
using ClearPlan.Mappers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using VMS.TPS.Common.Model.API;

namespace ClearPlan
{
    public class PQMSummaryCalculator : ViewModelBase
    {
        public PQMSummaryViewModel[] EvaluateTable(ConstraintViewModel selection, PlanningItemViewModel planningItem, ConstraintCatalog catalog)
        {
            var objectives = GetObjectives(selection);
            var structureSet = planningItem.PlanningItemStructureSet;
            var structures = structureSet == null ? new List<Structure>() : structureSet.Structures.Where(s => !s.IsEmpty).ToList();
            var candidates = structures.Select(s => new StructureCandidate
            {
                Id = s.Id, DicomType = s.DicomType,
                Codes = s.StructureCodeInfos == null ? new List<string>() : s.StructureCodeInfos
                    .Where(code => code != null && !string.IsNullOrWhiteSpace(code.Code)).Select(code => code.Code).ToList()
            }).ToList();
            bool hasValidDose = false;
            try
            {
                var plan = planningItem.PlanningItemObject as PlanSetup;
                var sum = planningItem.PlanningItemObject as PlanSum;
                hasValidDose = plan != null ? plan.IsDoseValid : sum != null && sum.IsDoseValid();
            }
            catch (Exception) { }
            for (int index = 0; index < objectives.Length; index++)
            {
                var objective = objectives[index];
                var definition = selection.Table.Constraints.Where(c => c != null).ElementAt(index);
                var match = StructureAliasResolver.Resolve(definition,
                    catalog == null ? new List<StructureDefinition>() : catalog.Structures, candidates);
                var structure = !match.IsAmbiguous && match.CandidateId != null
                    ? structures.SingleOrDefault(s => s.Id == match.CandidateId) : null;
                objective.Source = selection.SourceKind + ": " + selection.ConstraintName +
                    (string.IsNullOrWhiteSpace(definition.Source) || string.Equals(definition.Source.Trim(), "nan", StringComparison.OrdinalIgnoreCase)
                        ? "" : "; " + ClearPlan.Core.Review.ClinicalReviewValueMapper.SanitizeClinicalLabel(definition.Source, "Additional source configured"));
                objective.Comment = "Match: " + (structure == null ? match.Message : match.Rule + " -> " + structure.Id) +
                    "; selection: " + selection.SelectionReason + ". " + (definition.Comment ?? "");
                objective.ActivePlanningItem = planningItem;
                objective.EvaluationRequiresConfirmation = selection.RequiresConfirmation;
                if (selection.RequiresConfirmation)
                    objective.Comment += " Measurement uses the active plan dose without scaling. Goal assessment is pending confirmation of the constraint table and treatment scope.";
                objective.MappingDescription = structure == null ? "Unresolved or ambiguous" : match.Rule == StructureMatchRule.ExactRequestedName ? "Exact name"
                    : match.Rule == StructureMatchRule.CanonicalName ? "Canonical name" : match.Rule.ToString();
                objective.StructureList = StructureSetListViewModel.GetStructureList(structureSet);
                if (structure != null) { objective.StructureName = structure.Id; objective.StructureNameWithCode = structure.Id; }
                if (!hasValidDose || structure == null)
                {
                    objective.Achieved = !hasValidDose ? "Valid plan dose unavailable." : "Structure match unresolved or ambiguous.";
                    objective.isCalculated = false;
                    continue;
                }
                try { GetObjectiveProperties(objective, planningItem, structureSet, new StructureViewModel(structure)); }
                catch (Exception)
                {
                    // Keep the failed objective visible; never silently drop a constraint row.
                    objective.Achieved = "Metric unavailable; check objective expression, units and dose coverage.";
                    objective.Met = string.Empty; objective.isCalculated = false;
                }
            }
            return objectives;
        }

        public PQMSummaryViewModel[] GetObjectives(ConstraintViewModel constraintVM)
        {
            if (constraintVM == null || constraintVM.Table == null)
            {
                return new PQMSummaryViewModel[0];
            }

            PQMSummaryViewModel[] objectives = constraintVM.Table.Constraints
                .Where(constraint => constraint != null)
                .Select(PqmObjectiveViewModelMapper.Map)
                .ToArray();
            Console.WriteLine(
                "Constraint table {0} from {1} loaded. Objective number is {2}.",
                constraintVM.ConstraintId,
                constraintVM.SourceKind,
                objectives.Length);
            return objectives;
        }

      //  public ObservableCollection<Structure> FindMatchingStructures(ConstraintViewModel constraintVM, StructureSet structureSet)
      //  {
           // PQMSummaryViewModel[] m_objectives = GetObjectives(constraintVM);
           // var evalStructureList = new ObservableCollection<Structure>();
            //int i = 0;
           // foreach (var objective in m_objectives)
           // {
            //    Structure evalStructure = FindStructureFromAlias(structureSet, objective.TemplateId, objective.TemplateAliases, objective.TemplateCodes);
                //objective.Structure = evalStructure;
            //    evalStructureList.Add(evalStructure);
                //i++;
          //  }
         //   return evalStructureList;
       // }

        public PQMSummaryViewModel GetObjectiveProperties(PQMSummaryViewModel objective, PlanningItemViewModel planningItemVM, StructureSet structureSet, StructureViewModel evalStructure)
        {
            objective.ActivePlanningItem = planningItemVM;
            objective.StructureList = StructureSetListViewModel.GetStructureList(structureSet);
            // WPF SelectedItem must reference an actual item from this list. A separate
            // wrapper is not equal and can silently clear the selection during binding.
            var selected = evalStructure == null ? null : objective.StructureList.SingleOrDefault(
                item => string.Equals(item.StructureName, evalStructure.StructureName, StringComparison.OrdinalIgnoreCase));
            objective.NotifyPropertyChanged("StructureList");
            objective.Structure = selected;
            if (selected == null)
            {
                objective.Achieved = "Structure not found or empty.";
                objective.MappingDescription = "Unresolved or ambiguous";
                objective.NotifyPropertyChanged("Achieved");
                objective.NotifyPropertyChanged("MappingDescription");
            }
            return objective;
        }

        private static readonly Regex Whitespace = new Regex(@"\s+");
        public static string ReplaceWhitespace(string input, string replacement)
        {
            return Whitespace.Replace(input, replacement);
        }

        public Structure FindStructureFromAlias(StructureSet ss, PlanningItemViewModel planningItem, string ID, string[] aliases, string[] codes, string[] types)
        {
            if (ss == null) return null;
            var structures = ss.Structures.Where(structure => structure != null && !structure.IsEmpty).ToList();
            var definition = new StructureDefinition
            {
                CanonicalName = ID,
                Aliases = (aliases ?? new string[0]).ToList(),
                Codes = (codes ?? new string[0]).ToList(),
                Active = true
            };
            // All entry points share exact names, explicit aliases/codes and ambiguity handling.
            // The plan target and general DICOM types are not evidence for a requested organ.
            StructureMatch match = StructureAliasResolver.Resolve(
                new ConstraintDefinition { StructureName = ID },
                new[] { definition },
                structures.Select(structure => new StructureCandidate
                {
                    Id = structure.Id,
                    Codes = structure.StructureCodeInfos == null ? new List<string>() : structure.StructureCodeInfos
                        .Where(code => code != null && !string.IsNullOrWhiteSpace(code.Code)).Select(code => code.Code).ToList()
                }));
            return !match.IsAmbiguous && !string.IsNullOrWhiteSpace(match.CandidateId)
                ? structures.SingleOrDefault(structure => string.Equals(structure.Id, match.CandidateId, StringComparison.OrdinalIgnoreCase))
                : null;
        }

        void ConvertUnitToGy(ref string expression)
        {
            if (string.IsNullOrEmpty(expression)) return;
            expression = expression.Replace("cGy", "Gy");
        }

        void ConvertUnitTocGy(ref string expression)
        {
            if (string.IsNullOrEmpty(expression)) return;
            expression = expression.Replace("Gy", "cGy");
        }

        void ConvertValueToGy(ref string expression)
        {
            var resultString = Regex.Match(expression, @"\d+\p{P}\d+|\d+").Value;
            double newValue = double.NaN;
            if (double.TryParse(resultString, out newValue))
            {
                newValue = newValue / 100.0;
                expression = expression.Replace(resultString, newValue.ToString());
            }
        }

        void ConvertValueTocGy(ref string expression)
        {
            var resultString = Regex.Match(expression, @"\d+\p{P}\d+|\d+").Value;
            double newValue = double.NaN;
            if (double.TryParse(resultString, out newValue))
            {
                newValue = newValue * 100.0;
                expression = expression.Replace(resultString, newValue.ToString());
            }
        }
        public string CalculateMetric(StructureSet structureSet, StructureViewModel evalStructureVM, PlanningItemViewModel planningItem, string DVHObjective, string variation)
        {

            //start with a general regex that pulls out the metric type and the @ (evalunit) part.
            string pattern = @"^(?<type>[^\[\]]+)(\[(?<evalunit>[^\[\]]+)\])$";
            string minmaxmean_Pattern = @"^(M(in|ax|ean)|Volume)$";//check for Max or Min or Mean or Volume
            string d_at_v_pattern = @"^D(?<evalpt>\d+\p{P}\d+|\d+)(?<unit>(%|cc))$"; // matches D95%, D2cc
            string dc_at_v_pattern = @"^DC(?<evalpt>\d+)(?<unit>(%|cc))$"; // matches DC95%, DC700cc
            string v_at_d_pattern = @"^V(?<evalpt>\d+\p{P}\d+|\d+)(?<unit>(%|Gy|cGy))$"; // matches V98%, V40Gy
            string cv_at_d_pattern = @"^CV(?<evalpt>\d+(\.\d+)?)(?<unit>(%|Gy|cGy))$"; // matches CV98%, CV40Gy
                                                                               // Max[Gy] D95%[%] V98%[%] CV98%[%] D2cc[Gy] V40Gy[%]
            string cn_pattern = @"^CI(?<evalpt>\d+\p{P}\d+|\d+)(?<unit>(%|Gy|cGy))$"; //matches CN50%   
            string gi_pattern = @"^GI(?<evalpt>\d+\p{P}\d+|\d+)(?<unit>(%|Gy|cGy))$"; //matches GI50% 
            string hi_pattern = @"^HI(?<evalpt>\d+\p{P}\d+|\d+)(?<unit>(%|Gy|cGy))$"; //matches HI98%

            Structure evalStructure = evalStructureVM.Structure;
            if (evalStructure == null)
                return "Structure not found";

            //start with a general regex that pulls out the metric type and the [evalunit] part.
            var matches = Regex.Matches(DVHObjective, pattern);

            if (matches.Count != 1)
            {
                return string.Format("DVH Objective expression \"{0}\" is not a recognized expression type.", DVHObjective);
            }
            Match m = matches[0];
            Group type = m.Groups["type"];
            Group evalunit = m.Groups["evalunit"];
            Console.WriteLine("expression {0} => type = {1}, unit = {2}", DVHObjective, type.Value, evalunit.Value);

            // further decompose <type>
            var testMatch = Regex.Matches(type.Value, minmaxmean_Pattern);
            if (testMatch.Count != 1)
            {
                testMatch = Regex.Matches(type.Value, v_at_d_pattern);
                if (testMatch.Count != 1)
                {
                    testMatch = Regex.Matches(type.Value, d_at_v_pattern);
                    if (testMatch.Count != 1)
                    {
                        testMatch = Regex.Matches(type.Value, cv_at_d_pattern);
                        if (testMatch.Count != 1)
                        {
                            testMatch = Regex.Matches(type.Value, dc_at_v_pattern);
                            if (testMatch.Count != 1)
                            {
                                testMatch = Regex.Matches(type.Value, cn_pattern);
                                if (testMatch.Count != 1)
                                {
                                    testMatch = Regex.Matches(type.Value, gi_pattern);
                                    if (testMatch.Count != 1)
                                    {
                                        testMatch = Regex.Matches(type.Value, hi_pattern);
                                        if (testMatch.Count != 1)
                                        {
                                            return string.Format("DVH Objective expression \"{0}\" is not a recognized expression type.", DVHObjective);
                                        }
                                        else
                                        {
                                            // we have Homogenity Index pattern
                                            return PQMHomogenityIndex.GetHomogenityIndex(structureSet, planningItem, evalStructure, testMatch, evalunit);
                                        }
                                    }
                                    else
                                    {
                                        // we have Gradient Index pattern
                                        return (PQMGradientIndex.GetGradientIndex(structureSet, planningItem, evalStructure, testMatch, evalunit) == null? "Not evaluated" : PQMGradientIndex.GetGradientIndex(structureSet, planningItem, evalStructure, testMatch, evalunit));
                                    }
                                }
                                else
                                {
                                    // we have Conformation Number pattern
                                    return PQMConformationNumber.GetConformationNumber(structureSet, planningItem, evalStructure, testMatch, evalunit);
                                }

                            }
                            else
                            {
                                // we have Covered Dose at Volume pattern
                                return PQMCoveredDoseAtVolume.GetCoveredDoseAtVolume(structureSet, planningItem, evalStructure, testMatch, evalunit);
                            }
                        }
                        else
                        {
                            // we have Covered Volume at Dose pattern
                            return PQMCoveredVolumeAtDose.GetCoveredVolumeAtDose(structureSet, planningItem, evalStructure, testMatch, evalunit, variation);
                        }
                    }
                    else
                    {
                        // we have Dose at Volume pattern
                        return PQMDoseAtVolume.GetDoseAtVolume(structureSet, planningItem, evalStructure, testMatch, evalunit);
                    }
                }
                else
                {
                    // we have Volume at Dose pattern
                    return PQMVolumeAtDose.GetVolumeAtDose(structureSet, planningItem, evalStructure, testMatch, evalunit);
                }
            }
            else
            {
                // we have Min, Max, Mean, or Volume
                return PQMMinMaxMean.GetMinMaxMean(structureSet, planningItem, evalStructure, testMatch, evalunit, type);
            }
        }

        public string EvaluateMetric(string achieved, string goal, string variation)
        {
            return PqmNumericEvaluator.Evaluate(achieved, goal, variation);
        }
    }
}
