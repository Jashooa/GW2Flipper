namespace GW2Flipper.Extensions;

using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

public static class BitmapExtensions
{
    public static Bitmap BinarizeByColor(this Bitmap bitmap, Color color, double tolerance = 1.0d)
    {
        var bitmapData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
        var size = bitmapData.Stride * bitmapData.Height;
        var data = new byte[size];
        Marshal.Copy(bitmapData.Scan0, data, 0, size);

        for (var y = 0; y < bitmap.Height; ++y)
        {
            for (var x = 0; x < bitmap.Width; ++x)
            {
                var index = (y * bitmapData.Stride) + (x * Image.GetPixelFormatSize(bitmapData.PixelFormat) / 8);
                var pixelColor = Color.FromArgb(data[index + 2], data[index + 1], data[index]);
                if (Difference(pixelColor, color) < tolerance)
                {
                    data[index] = Color.White.B;
                    data[index + 1] = Color.White.G;
                    data[index + 2] = Color.White.R;
                }
                else
                {
                    data[index] = Color.Black.B;
                    data[index + 1] = Color.Black.G;
                    data[index + 2] = Color.Black.R;
                }
            }
        }

        Marshal.Copy(data, 0, bitmapData.Scan0, data.Length);

        bitmap.UnlockBits(bitmapData);

        return bitmap;
    }

    private static double Difference(Color c1, Color c2)
    {
        double diffR = Math.Abs(c1.R - c2.R);
        double diffG = Math.Abs(c1.G - c2.G);
        double diffB = Math.Abs(c1.B - c2.B);

        return 1.0 - (Math.Sqrt((diffR * diffR) + (diffG * diffG) + (diffB * diffB)) / 255.0);
    }

    public static Bitmap Invert(this Bitmap bitmap)
    {
        var bitmapData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
        var size = bitmapData.Stride * bitmapData.Height;
        var data = new byte[size];
        Marshal.Copy(bitmapData.Scan0, data, 0, size);

        for (var y = 0; y < bitmap.Height; ++y)
        {
            for (var x = 0; x < bitmap.Width; ++x)
            {
                var index = (y * bitmapData.Stride) + (x * Image.GetPixelFormatSize(bitmapData.PixelFormat) / 8);

                data[index] = (byte)(255 - data[index]);
                data[index + 1] = (byte)(255 - data[index + 1]);
                data[index + 2] = (byte)(255 - data[index + 2]);
            }
        }

        Marshal.Copy(data, 0, bitmapData.Scan0, data.Length);

        bitmap.UnlockBits(bitmapData);

        return bitmap;
    }

    public static Bitmap Contrast(this Bitmap bitmap, float value)
    {
        var bitmapData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
        var size = bitmapData.Stride * bitmapData.Height;
        var data = new byte[size];
        Marshal.Copy(bitmapData.Scan0, data, 0, size);

        for (var y = 0; y < bitmap.Height; ++y)
        {
            for (var x = 0; x < bitmap.Width; ++x)
            {
                var index = (y * bitmapData.Stride) + (x * Image.GetPixelFormatSize(bitmapData.PixelFormat) / 8);

                var r = data[index + 2];
                var g = data[index + 1];
                var b = data[index];

                var red = ((((r / 255.0f) - 0.5f) * value) + 0.5f) * 255.0f;
                var green = ((((g / 255.0f) - 0.5f) * value) + 0.5f) * 255.0f;
                var blue = ((((b / 255.0f) - 0.5f) * value) + 0.5f) * 255.0f;

                data[index + 2] = (byte)Math.Clamp(red, 0, 255);
                data[index + 1] = (byte)Math.Clamp(green, 0, 255);
                data[index] = (byte)Math.Clamp(blue, 0, 255);
            }
        }

        Marshal.Copy(data, 0, bitmapData.Scan0, data.Length);

        bitmap.UnlockBits(bitmapData);

        return bitmap;
    }

    public static Bitmap Binarize(this Bitmap bitmap, int v)
    {
        var bitmapData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
        var size = bitmapData.Stride * bitmapData.Height;
        var data = new byte[size];
        Marshal.Copy(bitmapData.Scan0, data, 0, size);

        for (var y = 0; y < bitmap.Height; ++y)
        {
            for (var x = 0; x < bitmap.Width; ++x)
            {
                var index = (y * bitmapData.Stride) + (x * Image.GetPixelFormatSize(bitmapData.PixelFormat) / 8);

                var r = data[index + 2];
                var g = data[index + 1];
                var b = data[index];

                data[index] = data[index + 1] = data[index + 2] = (byte)(((byte)((0.2125 * r) + (0.7154 * g) + (0.0721 * b)) >= v) ? 255 : 0);
            }
        }

        Marshal.Copy(data, 0, bitmapData.Scan0, data.Length);

        bitmap.UnlockBits(bitmapData);

        return bitmap;
    }

    public static Bitmap DrawBorder(this Bitmap bitmap, Color color, int borderSize = 10)
    {
        var newWidth = bitmap.Width + (borderSize * 2);
        var newHeight = bitmap.Height + (borderSize * 2);

        Image newImage = new Bitmap(newWidth, newHeight);
        using (var gfx = Graphics.FromImage(newImage))
        {
            using (Brush border = new SolidBrush(color))
            {
                gfx.FillRectangle(border, 0, 0, newWidth, newHeight);
            }
            gfx.DrawImage(bitmap, new Rectangle(borderSize, borderSize, bitmap.Width, bitmap.Height));
        }

        return (Bitmap)newImage;
    }
}
