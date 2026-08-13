using System;
using System.Linq;
using ClearPlan.Core.Constraints;

namespace ClearPlan.Core.Tests
{
    internal static class RefDbContractTests
    {
        public static int Run(string path)
        {
            ConstraintCatalog catalog = new RefDbJsonConstraintSource().Load(path);
            int fatalErrors = catalog.Issues.Count(issue => issue.IsFatal);

            Console.WriteLine("schema={0}", catalog.Schema);
            Console.WriteLine("tables.total={0}", catalog.Statistics.TotalTables);
            Console.WriteLine("tables.active={0}", catalog.Statistics.ActiveTables);
            Console.WriteLine("constraints.total={0}", catalog.Statistics.TotalConstraints);
            Console.WriteLine("structures.total={0}", catalog.Statistics.TotalStructures);
            Console.WriteLine("structures.active={0}", catalog.Statistics.ActiveStructures);
            Console.WriteLine("errors.fatal={0}", fatalErrors);

            return catalog.Schema == "RSAlign.local_refdb_constraints.v1" &&
                   catalog.Statistics.ActiveTables == 78 &&
                   catalog.Statistics.TotalConstraints == 2279 &&
                   catalog.Statistics.TotalStructures == 283 &&
                   catalog.Statistics.ActiveStructures == 185 &&
                   fatalErrors == 0
                ? 0
                : 1;
        }
    }
}
