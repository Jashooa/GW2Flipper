namespace GW2Flipper.Native;

using System.Runtime.InteropServices;

[StructLayout(LayoutKind.Sequential)]
internal struct RECT
{
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
}
