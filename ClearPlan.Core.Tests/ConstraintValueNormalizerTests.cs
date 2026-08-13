using System.Globalization;
using ClearPlan.Core.Constraints;

namespace ClearPlan.Core.Tests
{
    internal static class ConstraintValueNormalizerTests
    {
        public static void NormalizesComparators()
        {
            TestAssert.Equal("<=", ConstraintValueNormalizer.NormalizeComparator("≤"));
            TestAssert.Equal("<=", ConstraintValueNormalizer.NormalizeComparator("<="));
            TestAssert.Equal(">=", ConstraintValueNormalizer.NormalizeComparator("≥"));
            TestAssert.Equal(">=", ConstraintValueNormalizer.NormalizeComparator(">="));
            TestAssert.Equal("<", ConstraintValueNormalizer.NormalizeComparator("<"));
            TestAssert.Equal(string.Empty, ConstraintValueNormalizer.NormalizeComparator(null));
        }

        public static void NormalizesUnits()
        {
            TestAssert.Equal("cc", ConstraintValueNormalizer.NormalizeUnit("cm³"));
            TestAssert.Equal("cc", ConstraintValueNormalizer.NormalizeUnit(" cc "));
            TestAssert.Equal("Gy", ConstraintValueNormalizer.NormalizeUnit("gy"));
            TestAssert.Equal("%", ConstraintValueNormalizer.NormalizeUnit("%"));
            TestAssert.Equal("cm", ConstraintValueNormalizer.NormalizeUnit("cm (Länge)"));
        }

        public static void ParsesInvariantAndGermanDecimals()
        {
            CultureInfo previousCulture = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
                TestAssert.Equal(32.5m, ConstraintValueNormalizer.ParseNullableDecimal("32.5").Value);
                TestAssert.Equal(800.00m, ConstraintValueNormalizer.ParseNullableDecimal("800,00").Value);
                TestAssert.Equal(null, ConstraintValueNormalizer.ParseNullableDecimal(null));
                TestAssert.Equal(null, ConstraintValueNormalizer.ParseNullableDecimal(" "));
            }
            finally
            {
                CultureInfo.CurrentCulture = previousCulture;
            }
        }

        public static void NormalizesSupportedMetrics()
        {
            TestAssert.Equal("Dmean", ConstraintValueNormalizer.NormalizeMetric(" dmean "));
            TestAssert.Equal("Dmean", ConstraintValueNormalizer.NormalizeMetric("Mean"));
            TestAssert.Equal("Dmax", ConstraintValueNormalizer.NormalizeMetric("Max"));
            TestAssert.Equal("Dmin", ConstraintValueNormalizer.NormalizeMetric("Min"));
            TestAssert.Equal("D0.1cc", ConstraintValueNormalizer.NormalizeMetric("D0.1 cm³"));
            TestAssert.Equal("V20Gy", ConstraintValueNormalizer.NormalizeMetric("V20 gy"));
            TestAssert.Equal("D50%", ConstraintValueNormalizer.NormalizeMetric("D50 %"));
            TestAssert.Equal("CV21.5Gy", ConstraintValueNormalizer.NormalizeMetric("CV21.5 Gy"));
        }
    }
}
