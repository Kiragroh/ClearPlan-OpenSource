namespace ClearPlan.Core.Constraints
{
    public interface IConstraintCatalogSource
    {
        ConstraintCatalog Load(string path);
    }
}
