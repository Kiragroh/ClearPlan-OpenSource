using System;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace ClearPlan.Simulator
{
    public static class CapturePixelValidator
    {
        public static void AssertNonBlank(BitmapSource bitmap)
        {
            if (bitmap == null) throw new ArgumentNullException("bitmap");
            long pixelCount = (long)bitmap.PixelWidth * bitmap.PixelHeight;
            if (pixelCount < 1 || pixelCount > 20000000)
                throw new InvalidOperationException("Capture has invalid or excessive pixel dimensions.");
            BitmapSource source = bitmap;
            if (source.Format != PixelFormats.Bgra32 && source.Format != PixelFormats.Pbgra32)
                source = new FormatConvertedBitmap(source, PixelFormats.Bgra32, null, 0);
            int stride = checked(source.PixelWidth * 4);
            var pixels = new byte[checked(stride * source.PixelHeight)];
            source.CopyPixels(pixels, stride, 0);
            long opaque = 0;
            int minimum = 765, maximum = 0;
            for (int i = 0; i < pixels.Length; i += 4)
            {
                if (pixels[i + 3] < 250) continue;
                opaque++;
                int brightness = pixels[i] + pixels[i + 1] + pixels[i + 2];
                minimum = Math.Min(minimum, brightness);
                maximum = Math.Max(maximum, brightness);
            }
            // The captured application has an opaque background. Dimension-only PNG
            // checks miss a WPF no-display render that silently returns all-zero RGBA.
            if (opaque < pixelCount * 0.95 || maximum - minimum < 60)
                throw new InvalidOperationException(
                    "WPF capture is transparent, blank or almost entirely uniform. " +
                    "No PNG was written; verify automated off-screen rendering and layout readiness.");
        }
    }
}
