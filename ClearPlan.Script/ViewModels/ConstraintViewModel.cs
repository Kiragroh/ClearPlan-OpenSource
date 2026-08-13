using ClearPlan.Core.Constraints;

namespace ClearPlan
{
	public class ConstraintViewModel : ViewModelBase
    {
        public string ConstraintName { get; set; }
        public string ConstraintPath { get; set; }
        public string ConstraintId { get; set; }
        public string SourceKind { get; set; }
        public string SelectionReason { get; set; }
        public bool RequiresConfirmation { get; set; }
        public ConstraintTableDefinition Table { get; private set; }

        public ConstraintViewModel(
            ConstraintTableDefinition table,
            string constraintPath,
            string sourceKind)
        {
            Table = table;
            ConstraintId = table == null ? string.Empty : table.TableId;
            ConstraintName = table == null
                ? "Keine Constraint-Tabelle"
                : string.IsNullOrWhiteSpace(table.DisplayName)
                    ? table.TableId
                    : table.DisplayName;
            ConstraintPath = constraintPath ?? string.Empty;
            SourceKind = sourceKind ?? string.Empty;
        }
    }
}
