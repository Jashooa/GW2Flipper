namespace GW2Flipper.Native;

using System.Runtime.InteropServices;

internal static class Kernel32
{
    [DllImport("kernel32.dll")]
    public static extern uint GetCurrentThreadId();
}
