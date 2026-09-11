using System;

namespace ClearPlan.Core.Review
{
    public static class OrthogonalImageGeometry
    {
        // Conservative support: axis-aligned DICOM CT only. No oblique reformatting or guessed orientation.
        public static int AxisSign(double x, double y, double z, int expectedAxis)
        {
            double[] values = { x, y, z };
            if (expectedAxis < 0 || expectedAxis > 2) throw new ArgumentOutOfRangeException("expectedAxis");
            for (int i = 0; i < 3; i++)
            {
                double expectedMagnitude = i == expectedAxis ? 1 : 0;
                if (double.IsNaN(values[i]) || double.IsInfinity(values[i]) ||
                    Math.Abs(Math.Abs(values[i]) - expectedMagnitude) > 0.0001)
                    throw new ArgumentException("Oblique or nonstandard CT axes are not supported.");
            }
            return values[expectedAxis] > 0 ? 1 : -1;
        }

        public static int NearestIndex(double position, double origin, double spacing, int direction, int size)
        {
            if (size < 1 || spacing <= 0 || double.IsNaN(spacing) || double.IsInfinity(spacing) ||
                (direction != 1 && direction != -1)) throw new ArgumentException("Invalid image grid.");
            double index = (position - origin) / (spacing * direction);
            if (double.IsNaN(index) || double.IsInfinity(index) || index < 0 || index > size - 1)
                throw new ArgumentOutOfRangeException("position", "Isocenter lies outside the CT grid.");
            return (int)Math.Round(index, MidpointRounding.AwayFromZero);
        }

        public static byte WindowHu(double hu, double level, double width)
        {
            if (double.IsNaN(hu) || double.IsInfinity(hu) || width <= 0)
                throw new ArgumentException("Nonfinite HU or invalid CT window.");
            return (byte)Math.Round(Math.Max(0, Math.Min(255, (hu - (level - width / 2)) / width * 255)));
        }
    }
}
