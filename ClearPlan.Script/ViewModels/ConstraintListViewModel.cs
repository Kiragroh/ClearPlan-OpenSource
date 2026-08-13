using System.Collections.ObjectModel;
using System.Linq;
using ClearPlan.Core.Constraints;

namespace ClearPlan
{
    class ConstraintListViewModel : ViewModelBase
    {
        public static ObservableCollection<ConstraintViewModel> GetConstraintList(
            ConstraintCatalogLoadResult loadResult)
        {
            var constraintComboBoxList = new ObservableCollection<ConstraintViewModel>();
            if (loadResult == null || !loadResult.IsUsable)
            {
                return constraintComboBoxList;
            }

            foreach (ConstraintTableDefinition table in loadResult.Catalog.Tables
                .Where(item => item != null && item.Active)
                .OrderBy(item => item.IsPlanSum ? 1 : 0)
                .ThenBy(item => item.FractionCountMinimum ?? int.MaxValue)
                .ThenBy(item => item.DisplayName)
                .ThenBy(item => item.TableId))
            {
                constraintComboBoxList.Add(new ConstraintViewModel(
                    table,
                    loadResult.ActivePath,
                    loadResult.ActiveSource));
            }

            return constraintComboBoxList;
        }
    }
}
