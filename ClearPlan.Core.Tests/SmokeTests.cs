using ClearPlan.Core.Constraints;

namespace ClearPlan.Core.Tests
{
    internal static class SmokeTests
    {
        public static void ConstraintSourceModeIsStable()
        {
            TestAssert.Equal("Automatic", ConstraintSourceMode.Automatic.ToString());
        }
    }
}
