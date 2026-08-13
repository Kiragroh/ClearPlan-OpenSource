using System;
using System.IO;
using System.Reflection;
using ClearPlan.Core.Settings;

namespace ClearPlan.Core.Tests
{
    internal static class PathFailureClassifierTests
    {
        public static void RejectsCollectionIndexErrors()
        {
            TestAssert.False(
                PathFailureClassifier.IsPathFailure(
                    new ArgumentOutOfRangeException("index")));
        }

        public static void RecognizesWrappedIoErrors()
        {
            TestAssert.True(
                PathFailureClassifier.IsPathFailure(
                    new TargetInvocationException(
                        new IOException("Network path unavailable."))));
        }
    }
}
