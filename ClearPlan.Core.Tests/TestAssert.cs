using System;
using System.Collections.Generic;

namespace ClearPlan.Core.Tests
{
    internal static class TestAssert
    {
        public static void Equal<T>(T expected, T actual, string message = null)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
            {
                throw new InvalidOperationException(
                    message ?? string.Format("Expected <{0}> but found <{1}>.", expected, actual));
            }
        }

        public static void True(bool condition, string message = null)
        {
            if (!condition)
            {
                throw new InvalidOperationException(message ?? "Expected condition to be true.");
            }
        }

        public static void False(bool condition, string message = null)
        {
            if (condition)
            {
                throw new InvalidOperationException(message ?? "Expected condition to be false.");
            }
        }

        public static void NotNull(object value, string message = null)
        {
            if (value == null)
            {
                throw new InvalidOperationException(message ?? "Expected a non-null value.");
            }
        }

        public static TException Throws<TException>(Action action, string message = null)
            where TException : Exception
        {
            try
            {
                action();
            }
            catch (TException exception)
            {
                return exception;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    message ?? string.Format(
                        "Expected {0}, but caught {1}.",
                        typeof(TException).Name,
                        exception.GetType().Name));
            }

            throw new InvalidOperationException(
                message ?? string.Format("Expected {0} to be thrown.", typeof(TException).Name));
        }
    }
}
