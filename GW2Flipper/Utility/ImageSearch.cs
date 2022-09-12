namespace GW2Flipper.Utility;

using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

using global::GW2Flipper.Native;

using NLog;

internal static class ImageSearch
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

    public static Bitmap? CaptureFullWindow(Process process)
    {
        if (process == null)
        {
            return null;
        }

        _ = User32.GetClientRect(process.MainWindowHandle, out var clientRect);
        var width = clientRect.Right - clientRect.Left;
        var height = clientRect.Bottom - clientRect.Top;

        return CaptureWindow(process, 0, 0, width, height);
    }

    public static Bitmap? CaptureWindow(Process process, int x, int y, int width, int height)
    {
        if (process == null)
        {
            return null;
        }

        Input.EnsureForegroundWindow(process);

        var clientPoint = new Point(x, y);
        _ = User32.ClientToScreen(process.MainWindowHandle, ref clientPoint);

        var hdcFrom = User32.GetDC(IntPtr.Zero);
        var hdcTo = Gdi32.CreateCompatibleDC(hdcFrom);
        var hBitmap = Gdi32.CreateCompatibleBitmap(hdcFrom, width, height);

        Bitmap? image = null;

        if (hBitmap != IntPtr.Zero)
        {
            var hLocalBitmap = Gdi32.SelectObject(hdcTo, hBitmap);
            _ = Gdi32.BitBlt(hdcTo, 0, 0, width, height, hdcFrom, clientPoint.X, clientPoint.Y, Gdi32.TernaryRasterOperations.SRCCOPY);
            _ = Gdi32.SelectObject(hdcTo, hLocalBitmap);

            image = Image.FromHbitmap(hBitmap);
            _ = Gdi32.DeleteObject(hBitmap);
        }

        _ = Gdi32.DeleteDC(hdcTo);
        _ = User32.ReleaseDC(process.MainWindowHandle, hdcFrom);

        return image;
    }

    public static Point? FindImageInFullWindow(Process process, Bitmap image, double tolerance = 1.0)
    {
        var windowImage = CaptureFullWindow(process);

        return Find(windowImage!, image, tolerance);
    }

    public static Point? FindImageInWindow(Process process, Bitmap image, int x, int y, int width, int height, double tolerance = 1.0)
    {
        var windowImage = CaptureWindow(process, x, y, width, height);

        var point = Find(windowImage!, image, tolerance);

        if (point != null)
        {
            point = Point.Add(point.Value, new Size(x, y));
        }

        return point;
    }

    public static Point? Find(Bitmap haystack, Bitmap needle, double tolerance = 1.0)
    {
        if (haystack == null || needle == null)
        {
            return null;
        }

        if (haystack.Width < needle.Width || haystack.Height < needle.Height)
        {
            return null;
        }

        var haystackArray = GetPixelArray(haystack);
        var needleArray = GetPixelArray(needle);

        foreach (var firstLineMatchPoint in FindMatch(haystackArray.Take(haystack.Height - needle.Height + 1), needleArray[0], tolerance))
        {
            if (IsNeedlePresentAtLocation(haystackArray, needleArray, firstLineMatchPoint, 1, tolerance))
            {
                return firstLineMatchPoint;
            }
        }

        return null;
    }

    private static int[][] GetPixelArray(Bitmap bitmap)
    {
        var result = new int[bitmap.Height][];
        var bitmapData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

        for (var y = 0; y < bitmap.Height; ++y)
        {
            result[y] = new int[bitmap.Width];
            Marshal.Copy(bitmapData.Scan0 + (y * bitmapData.Stride), result[y], 0, result[y].Length);
        }

        bitmap.UnlockBits(bitmapData);

        return result;
    }

    private static IEnumerable<Point> FindMatch(IEnumerable<int[]> haystackLines, int[] needleLine, double tolerance = 1.0)
    {
        var y = 0;
        foreach (var haystackLine in haystackLines)
        {
            for (int x = 0, n = haystackLine.Length - needleLine.Length + 1; x < n; ++x)
            {
                if (ContainSameElements(haystackLine, x, needleLine, 0, needleLine.Length, tolerance))
                {
                    yield return new Point(x, y);
                }
            }

            y++;
        }
    }

    private static bool ContainSameElements(int[] first, int firstStart, int[] second, int secondStart, int length, double tolerance = 1.0)
    {
        for (var i = 0; i < length; ++i)
        {
            if (tolerance == 1.0)
            {
                if (first[i + firstStart] != second[i + secondStart])
                {
                    return false;
                }
            }
            else if (GetMetric(first[i + firstStart], second[i + secondStart]) < tolerance)
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsNeedlePresentAtLocation(int[][] haystack, int[][] needle, Point point, int alreadyVerified, double tolerance = 1.0)
    {
        // we already know that "alreadyVerified" lines already match, so skip them
        for (var y = alreadyVerified; y < needle.Length; ++y)
        {
            if (!ContainSameElements(haystack[y + point.Y], point.X, needle[y], 0, needle[y].Length, tolerance))
            {
                return false;
            }
        }

        return true;
    }

    private static double GetMetric(int first, int second)
    {
        // double a = (byte)(first >> 24) - (byte)(second >> 24);
        double r = (byte)(first >> 16) - (byte)(second >> 16);
        double g = (byte)(first >> 8) - (byte)(second >> 8);
        double b = (byte)(first >> 0) - (byte)(second >> 0);

        // return 1.0 - Math.Sqrt((a * a) + (r * r) + (g * g) + (b * b)) / 255;
        return 1.0 - (Math.Sqrt((r * r) + (g * g) + (b * b)) / 255);
    }
}
