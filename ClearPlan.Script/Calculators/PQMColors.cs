using System;
using ClearPlan.Core.Constraints;
using System.Text.RegularExpressions;
using System.Windows.Media;
using VMS.TPS.Common.Model.API;

namespace ClearPlan.Calculators
{
    public class PQMColors
    {
        public static Tuple<SolidColorBrush, double> GetAchievedColor(Structure structure, string goal, string DVHObjective, string Achieved)
        {
            var achievedColor = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFFFFF"));
            double achievedDouble;
            double achievedRatio = 0;
            double goalDouble;
            string comparator;
            if (structure == null || !PqmNumericEvaluator.TryParseAchieved(Achieved, out achievedDouble) ||
                !PqmNumericEvaluator.TryParseGoal(goal, out comparator, out goalDouble) ||
                !PqmNumericEvaluator.TryGetGoalRatio(Achieved, goal, out achievedRatio))
            {
                achievedColor = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFFFFF")); //white
                achievedDouble = 0;
                return Tuple.Create(achievedColor, achievedRatio);
            }          
            if (goal.Contains("<"))  //D at V, V at D, serial tissue
            {
                achievedColor = GetNormalTissueSolidColorBrush(achievedRatio);
            }
            if (goal.Contains(">") && (DVHObjective ?? string.Empty).Contains("CV"))  //CV parallel tissue
            {
                var outputUnit = Regex.Match(DVHObjective, @"\[(?<unit>%|cc)\]$").Groups["unit"].Value;
                if (!PqmNumericEvaluator.TryGetComplementGoalRatio(Achieved, goal, structure.Volume, outputUnit, out achievedRatio))
                    return Tuple.Create(achievedColor, 0.0);
                achievedColor = GetNormalTissueSolidColorBrush(achievedRatio);
            }

            else if (goal.ToString().Contains(">")) //Target
            {
                achievedColor = GetTargetSolidColorBrush(achievedRatio);
            }
            return Tuple.Create(achievedColor, achievedRatio);
        }

        private static SolidColorBrush GetNormalTissueSolidColorBrush(double achievedRatio)
        {
            var achievedColor = new SolidColorBrush();
            if (achievedRatio >= 100)
            {
                achievedRatio = 100;
                achievedColor = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFE07979")); //red
            }
            if (achievedRatio >= 75 && achievedRatio < 100)
                achievedColor = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFEEA862")); //orange;
            if (achievedRatio >= 50 && achievedRatio < 75)
                achievedColor = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFF0F3A4")); //yellow;
            if (achievedRatio >= 25 && achievedRatio < 50)
                achievedColor = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFA7F3A4")); //light green
            if (achievedRatio >= 0 && achievedRatio < 25)
                achievedColor = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFA7F3A4")); //light green
            if (achievedRatio < 0)
                achievedColor = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFA7F3A4")); //light green
            return achievedColor;
        }

        private static SolidColorBrush GetTargetSolidColorBrush(double achievedRatio)
        {
            var achievedColor = new SolidColorBrush();
            if (achievedRatio >= 100)
            {
                achievedRatio = 100;
                achievedColor = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFA7F3A4")); //light green
            }
            if (achievedRatio >= 90 && achievedRatio < 100)
                achievedColor = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFA7F3A4")); //light green
            if (achievedRatio >= 50 && achievedRatio < 90)
                achievedColor = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFEEA862")); //orange;
            if (achievedRatio < 50)
                achievedColor = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFE07979")); //red;
            return achievedColor;
        }
    }
}
