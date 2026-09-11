using System;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ClearPlan.Simulator;

namespace ClearPlan.Core.Tests
{
    internal static class CapturePixelValidatorTests
    {
        public static void RejectsTransparentAndSolidCaptures()
        {
            TestAssert.Throws<InvalidOperationException>(() => CapturePixelValidator.AssertNonBlank(Bitmap(new byte[400])));
            var black = new byte[400];
            for (int i = 3; i < black.Length; i += 4) black[i] = 255;
            TestAssert.Throws<InvalidOperationException>(() => CapturePixelValidator.AssertNonBlank(Bitmap(black)));
            var sparse = new byte[400];
            for (int i = 0; i < 4; i++) sparse[i] = 255;
            TestAssert.Throws<InvalidOperationException>(() => CapturePixelValidator.AssertNonBlank(Bitmap(sparse)));
        }

        public static void AcceptsOpaqueVariedCapturePixels()
        {
            var pixels = new byte[400];
            for (int i = 0; i < pixels.Length; i += 4)
            {
                byte value = (byte)(i < 200 ? 32 : 240);
                pixels[i] = pixels[i + 1] = pixels[i + 2] = value;
                pixels[i + 3] = 255;
            }
            CapturePixelValidator.AssertNonBlank(Bitmap(pixels));
        }

        private static BitmapSource Bitmap(byte[] pixels)
        {
            return BitmapSource.Create(10, 10, 96, 96, PixelFormats.Bgra32, null, pixels, 40);
        }
    }
}
