using System;

namespace ClearPlan.Core.Review
{
    public sealed class ReviewWindowSize
    {
        public ReviewWindowSize(double width, double height)
        {
            Width = width;
            Height = height;
        }

        public double Width { get; private set; }

        public double Height { get; private set; }
    }

    public static class ReviewWindowSizePolicy
    {
        public const double MinimumWidth = 1180.0;
        public const double MinimumHeight = 720.0;
        public const double MaximumInitialWidth = 1680.0;
        public const double WidthFraction = 0.92;
        public const double HeightFraction = 0.88;

        public static ReviewWindowSize Calculate(
            double availableWidth,
            double availableHeight)
        {
            ValidateAvailableDimension(
                availableWidth,
                "availableWidth");
            ValidateAvailableDimension(
                availableHeight,
                "availableHeight");

            double width = Math.Min(
                MaximumInitialWidth,
                Math.Max(
                    MinimumWidth,
                    availableWidth * WidthFraction));
            double height = Math.Max(
                MinimumHeight,
                availableHeight * HeightFraction);

            return new ReviewWindowSize(width, height);
        }

        private static void ValidateAvailableDimension(
            double value,
            string parameterName)
        {
            if (double.IsNaN(value) ||
                double.IsInfinity(value) ||
                value <= 0.0)
            {
                throw new ArgumentOutOfRangeException(
                    parameterName,
                    "The available work-area dimension must be finite and positive.");
            }
        }
    }
}
