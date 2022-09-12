namespace GW2Flipper.Native;

using System.Runtime.InteropServices;

internal static class Gdi32
{
    public enum TernaryRasterOperations : uint
    {
        BLACKNESS = 0x00000042,
        NOTSRCERASE = 0x001100A6,
        NOTSRCCOPY = 0x00330008,
        SRCERASE = 0x00440328,
        DSTINVERT = 0x00550009,
        PATINVERT = 0x005A0049,
        SRCINVERT = 0x00660046,
        SRCAND = 0x008800C6,
        MERGEPAINT = 0x00BB0226,
        MERGECOPY = 0x00C000CA,
        SRCCOPY = 0x00CC0020,
        SRCPAINT = 0x00EE0086,
        PATCOPY = 0x00F00021,
        PATPAINT = 0x00FB0A09,
        WHITENESS = 0x00FF0062,
        CAPTUREBLT = 0x40000000,
    }

    [DllImport("gdi32.dll", EntryPoint = "BitBlt", SetLastError = true)]
    public static extern bool BitBlt(IntPtr hdc, int x, int y, int cx, int cy, IntPtr hdcSrc, int x1, int y1, TernaryRasterOperations rop);

    [DllImport("gdi32.dll", EntryPoint = "CreateCompatibleBitmap")]
    public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int cx, int cy);

    [DllImport("gdi32.dll", EntryPoint = "CreateCompatibleDC", SetLastError = true)]
    public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

    [DllImport("gdi32.dll", EntryPoint = "DeleteDC")]
    public static extern bool DeleteDC(IntPtr hdc);

    [DllImport("gdi32.dll", EntryPoint = "DeleteObject")]
    public static extern bool DeleteObject(IntPtr ho);

    [DllImport("gdi32.dll", EntryPoint = "SelectObject")]
    public static extern IntPtr SelectObject(IntPtr hDC, IntPtr h);
}
