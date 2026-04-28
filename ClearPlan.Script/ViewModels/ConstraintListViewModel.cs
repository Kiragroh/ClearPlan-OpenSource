using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace ClearPlan
{
    class ConstraintListViewModel : ViewModelBase
    {
        public static ObservableCollection<ConstraintViewModel> GetConstraintList(string constraintDir)
        {
            var constraintComboBoxList = new ObservableCollection<ConstraintViewModel>();
            if (!Directory.Exists(constraintDir))
            {
                return constraintComboBoxList;
            }

            var files = Directory.EnumerateFiles(constraintDir, "*.csv")
                .OrderBy(file => Path.GetFileName(file).StartsWith("Starter_", StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                .ThenBy(file => Path.GetFileName(file), StringComparer.OrdinalIgnoreCase);

            foreach (string file in files)
            {
                constraintComboBoxList.Add(new ConstraintViewModel(file));
            }

            return constraintComboBoxList;
        }
    }
}
