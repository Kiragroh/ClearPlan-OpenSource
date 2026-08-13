using System;
using System.IO;
using System.Reflection;
using System.Security;

namespace ClearPlan.Core.Settings
{
    public static class PathFailureClassifier
    {
        public static bool IsPathFailure(Exception exception)
        {
            Exception current = Unwrap(exception);
            return current is IOException ||
                   current is UnauthorizedAccessException ||
                   current is NotSupportedException ||
                   current is SecurityException;
        }

        public static Exception Unwrap(Exception exception)
        {
            Exception current = exception ??
                                new InvalidOperationException(
                                    "Unknown application failure.");
            while ((current is TargetInvocationException ||
                    current is AggregateException) &&
                   current.InnerException != null)
            {
                current = current.InnerException;
            }

            return current;
        }
    }
}
