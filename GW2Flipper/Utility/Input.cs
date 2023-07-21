namespace GW2Flipper.Utility;

using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

using global::GW2Flipper.Native;

using NLog;

using TextCopy;

internal static class Input
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

    public enum MouseButton
    {
        LeftButton = 0,
        MiddleButton = 1,
        RightButton = 2,
    }

    public static void MouseMoveBackground(Process process, int x, int y, int delay = 25)
    {
        _ = User32.PostMessage(process.MainWindowHandle, WindowMessage.WM_MOUSEMOVE, IntPtr.Zero, GetLParam(x, y));
        Thread.Sleep(delay);
    }

    public static void MouseClickBackground(Process process, VirtualKeyCode key, int x, int y, int delay = 25)
    {
        MouseMoveBackground(process, x, y);
        Thread.Sleep(delay);

        switch (key)
        {
            case VirtualKeyCode.LBUTTON:
                _ = User32.PostMessage(process.MainWindowHandle, WindowMessage.WM_LBUTTONDOWN, (IntPtr)key, GetLParam(x, y));
                Thread.Sleep(delay);
                _ = User32.PostMessage(process.MainWindowHandle, WindowMessage.WM_LBUTTONUP, IntPtr.Zero, GetLParam(x, y));
                break;

            case VirtualKeyCode.RBUTTON:
                _ = User32.PostMessage(process.MainWindowHandle, WindowMessage.WM_RBUTTONDOWN, (IntPtr)key, GetLParam(x, y));
                Thread.Sleep(delay);
                _ = User32.PostMessage(process.MainWindowHandle, WindowMessage.WM_RBUTTONUP, IntPtr.Zero, GetLParam(x, y));
                break;

            case VirtualKeyCode.MBUTTON:
                _ = User32.PostMessage(process.MainWindowHandle, WindowMessage.WM_MBUTTONDOWN, (IntPtr)key, GetLParam(x, y));
                Thread.Sleep(delay);
                _ = User32.PostMessage(process.MainWindowHandle, WindowMessage.WM_MBUTTONUP, IntPtr.Zero, GetLParam(x, y));
                break;
        }

        Thread.Sleep(delay);
    }

    public static void SendKeyBackground(Process process, VirtualKeyCode key, int delay = 100)
    {
        _ = User32.PostMessage(process.MainWindowHandle, WindowMessage.WM_KEYDOWN, (IntPtr)key, GetLParam(1, key, 0, 0, 0, 0));
        Thread.Sleep(delay);
        _ = User32.PostMessage(process.MainWindowHandle, WindowMessage.WM_KEYUP, (IntPtr)key, GetLParam(1, key, 0, 0, 1, 1));
        Thread.Sleep(delay);
    }

    public static void SendKeysBackground(Process process, string keys, int delay = 25)
    {
        foreach (var key in keys)
        {
            _ = User32.PostMessage(process.MainWindowHandle, WindowMessage.WM_CHAR, (IntPtr)key, IntPtr.Zero);
            Thread.Sleep(delay);
        }
    }

    public static void MouseMove(Process process, int x, int y)
    {
        EnsureForegroundWindow(process);

        var point = new Point(x, y);
        var success = User32.ClientToScreen(process.MainWindowHandle, ref point);
        if (!success)
        {
            Logger.Error("ClientToScreen failed.");
            throw new Exception("ClientToScreen failed.");
        }

        GetNormalizedPoint(ref point);

        var input = new INPUT
        {
            Type = (uint)InputType.Mouse,
            Data =
            {
                Mouse = new MOUSEINPUT
                {
                    Flags = (uint)(MouseFlag.Move | MouseFlag.Absolute | MouseFlag.VirtualDesk),
                    X = point.X,
                    Y = point.Y,
                    Time = 0,
                    ExtraInfo = IntPtr.Zero,
                },
            },
        };

        SendInput(input);
    }

    public static void MouseButtonDown(Process process, MouseButton button)
    {
        EnsureForegroundWindow(process);

        var input = new INPUT
        {
            Type = (uint)InputType.Mouse,
            Data =
            {
                Mouse = new MOUSEINPUT
                {
                    Flags = (uint)ToMouseButtonDownFlag(button),
                    Time = 0,
                    ExtraInfo = IntPtr.Zero,
                },
            },
        };

        SendInput(input);
    }

    public static void MouseButtonUp(Process process, MouseButton button)
    {
        EnsureForegroundWindow(process);

        var input = new INPUT
        {
            Type = (uint)InputType.Mouse,
            Data =
            {
                Mouse = new MOUSEINPUT
                {
                    Flags = (uint)ToMouseButtonUpFlag(button),
                    Time = 0,
                    ExtraInfo = IntPtr.Zero,
                },
            },
        };

        SendInput(input);
    }

    public static void MouseClick(Process process, MouseButton button, int delay = 100)
    {
        MouseButtonDown(process, button);
        Thread.Sleep(delay);
        MouseButtonUp(process, button);
        Thread.Sleep(delay);
    }

    public static void MouseDoubleClick(Process process, MouseButton button, int delay = 100)
    {
        MouseButtonDown(process, button);
        Thread.Sleep(delay);
        MouseButtonUp(process, button);
        Thread.Sleep(delay);
        MouseButtonDown(process, button);
        Thread.Sleep(delay);
        MouseButtonUp(process, button);
        Thread.Sleep(delay);
    }

    public static void MouseMoveAndClick(Process process, MouseButton button, int x, int y, int delay = 100)
    {
        MouseMove(process, x, y);
        Thread.Sleep(delay);
        MouseClick(process, button, delay);
    }

    public static void MouseMoveAndDoubleClick(Process process, MouseButton button, int x, int y, int delay = 100)
    {
        MouseMove(process, x, y);
        Thread.Sleep(delay);
        MouseDoubleClick(process, button, delay);
    }

    public static void KeyDown(Process process, VirtualKeyCode keyCode)
    {
        EnsureForegroundWindow(process);

        var input = new INPUT
        {
            Type = (uint)InputType.Keyboard,
            Data =
            {
                Keyboard = new KEYBDINPUT
                {
                    KeyCode = (ushort)keyCode,
                    Scan = (ushort)(GetScanCode(keyCode) & 0xFFU),
                    Flags = IsExtendedKey(keyCode) ? (uint)KeyboardFlag.ExtendedKey : (uint)KeyboardFlag.None,
                    Time = 0,
                    ExtraInfo = IntPtr.Zero,
                },
            },
        };

        SendInput(input);
    }

    public static void KeyUp(Process process, VirtualKeyCode keyCode)
    {
        EnsureForegroundWindow(process);

        var input = new INPUT
        {
            Type = (uint)InputType.Keyboard,
            Data =
            {
                Keyboard = new KEYBDINPUT
                {
                    KeyCode = (ushort)keyCode,
                    Scan = (ushort)(GetScanCode(keyCode) & 0xFFU),
                    Flags = (uint)KeyboardFlag.KeyUp | (IsExtendedKey(keyCode) ? (uint)KeyboardFlag.ExtendedKey : (uint)KeyboardFlag.None),
                    Time = 0,
                    ExtraInfo = IntPtr.Zero,
                },
            },
        };

        SendInput(input);
    }

    public static void KeyPress(Process process, VirtualKeyCode keyCode, int delay = 100)
    {
        KeyDown(process, keyCode);
        Thread.Sleep(delay);
        KeyUp(process, keyCode);
        Thread.Sleep(delay);
    }

    public static void KeyPressWithModifier(Process process, VirtualKeyCode keyCode, bool alt, bool control, bool shift, int delay = 100)
    {
        if (alt)
        {
            KeyDown(process, VirtualKeyCode.MENU);
            Thread.Sleep(delay);
        }

        if (control)
        {
            KeyDown(process, VirtualKeyCode.CONTROL);
            Thread.Sleep(delay);
        }

        if (shift)
        {
            KeyDown(process, VirtualKeyCode.SHIFT);
            Thread.Sleep(delay);
        }

        KeyPress(process, keyCode, delay);

        if (shift)
        {
            KeyUp(process, VirtualKeyCode.SHIFT);
            Thread.Sleep(delay);
        }

        if (control)
        {
            KeyUp(process, VirtualKeyCode.CONTROL);
            Thread.Sleep(delay);
        }

        if (alt)
        {
            KeyUp(process, VirtualKeyCode.MENU);
            Thread.Sleep(delay);
        }
    }

    public static void KeyCharDown(Process process, char character)
    {
        EnsureForegroundWindow(process);

        var input = new INPUT
        {
            Type = (uint)InputType.Keyboard,
            Data =
            {
                Keyboard = new KEYBDINPUT
                {
                    KeyCode = 0,
                    Scan = character,
                    Flags = (uint)KeyboardFlag.Unicode,
                    Time = 0,
                    ExtraInfo = IntPtr.Zero,
                },
            },
        };

        if ((character & 0xFF00) == 0xE000)
        {
            input.Data.Keyboard.Flags |= (uint)KeyboardFlag.ExtendedKey;
        }

        SendInput(input);
    }

    public static void KeyCharUp(Process process, char character)
    {
        EnsureForegroundWindow(process);

        var input = new INPUT
        {
            Type = (uint)InputType.Keyboard,
            Data =
            {
                Keyboard = new KEYBDINPUT
                {
                    KeyCode = 0,
                    Scan = character,
                    Flags = (uint)KeyboardFlag.KeyUp | (uint)KeyboardFlag.Unicode,
                    Time = 0,
                    ExtraInfo = IntPtr.Zero,
                },
            },
        };

        if ((character & 0xFF00) == 0xE000)
        {
            input.Data.Keyboard.Flags |= (uint)KeyboardFlag.ExtendedKey;
        }

        SendInput(input);
    }

    public static void KeyCharPress(Process process, char character, int delay = 100)
    {
        KeyCharDown(process, character);
        Thread.Sleep(delay);
        KeyCharUp(process, character);
        Thread.Sleep(delay);
    }

    public static void KeyStringSend(Process process, string text, int delay = 100)
    {
        foreach (var character in text)
        {
            KeyCharPress(process, character, delay);
        }
    }

    public static void KeyStringSendClipboard(Process process, string text)
    {
        ClipboardService.SetText(text);
        Thread.Sleep(50);
        KeyPressWithModifier(process, VirtualKeyCode.VK_V, false, true, false);
        ClipboardService.SetText(string.Empty);
    }

    public static string? GetSelectedText(Process process)
    {
        KeyPressWithModifier(process, VirtualKeyCode.VK_C, false, true, false);
        Thread.Sleep(50);
        var text = ClipboardService.GetText();
        ClipboardService.SetText(string.Empty);

        return text;
    }

    public static void SendInput(INPUT input)
    {
        INPUT[] inputs = { input };
        _ = User32.SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        /*var success = User32.SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        if (success != inputs.Length)
        {
            throw new Exception("Some SendInput inputs were not sent successfully.");
        }*/
    }

    public static void EnsureForegroundWindow(Process process)
    {
        if (User32.GetForegroundWindow() != process.MainWindowHandle)
        {
            _ = User32.SetForegroundWindow(process.MainWindowHandle);
        }
    }

    private static IntPtr GetLParam(int x, int y) => (IntPtr)((y << 16) | (x & 0xFFFF));

    private static IntPtr GetLParam(uint repeatCount, VirtualKeyCode key, uint extended, uint context, uint previousState, uint transition)
    {
        var scanCode = GetScanCode(key);

        var lParam = repeatCount
            | (scanCode << 16)
            | (extended << 24)
            | (context << 29)
            | (previousState << 30)
            | (transition << 31);

        return (IntPtr)lParam;
    }

    private static void GetNormalizedPoint(ref Point point)
    {
        point.X = (point.X * 65526 / User32.GetSystemMetrics(0)) + (point.X < 0 ? -1 : 1);
        point.Y = (point.Y * 65526 / User32.GetSystemMetrics(1)) + (point.Y < 0 ? -1 : 1);
    }

    private static MouseFlag ToMouseButtonDownFlag(MouseButton button) => button switch
    {
        MouseButton.LeftButton => MouseFlag.LeftDown,
        MouseButton.MiddleButton => MouseFlag.MiddleDown,
        MouseButton.RightButton => MouseFlag.RightDown,
        _ => MouseFlag.LeftDown,
    };

    private static MouseFlag ToMouseButtonUpFlag(MouseButton button) => button switch
    {
        MouseButton.LeftButton => MouseFlag.LeftUp,
        MouseButton.MiddleButton => MouseFlag.MiddleUp,
        MouseButton.RightButton => MouseFlag.RightUp,
        _ => MouseFlag.LeftUp,
    };

    private static bool IsExtendedKey(VirtualKeyCode keyCode) => keyCode is VirtualKeyCode.MENU
        or VirtualKeyCode.LMENU
        or VirtualKeyCode.RMENU
        or VirtualKeyCode.CONTROL
        or VirtualKeyCode.RCONTROL
        or VirtualKeyCode.INSERT
        or VirtualKeyCode.DELETE
        or VirtualKeyCode.HOME
        or VirtualKeyCode.END
        or VirtualKeyCode.PRIOR
        or VirtualKeyCode.NEXT
        or VirtualKeyCode.RIGHT
        or VirtualKeyCode.UP
        or VirtualKeyCode.LEFT
        or VirtualKeyCode.DOWN
        or VirtualKeyCode.NUMLOCK
        or VirtualKeyCode.CANCEL
        or VirtualKeyCode.SNAPSHOT
        or VirtualKeyCode.DIVIDE;

    private static uint GetScanCode(VirtualKeyCode key) => User32.MapVirtualKey((uint)key, MapType.MAPVK_VK_TO_VSC);

    private static string GetCharFromKey(VirtualKeyCode key, bool alt, bool control, bool shift)
    {
        var buffer = new StringBuilder(256);
        var keyboardState = new byte[256];
        if (alt)
        {
            keyboardState[(int)VirtualKeyCode.SHIFT] = 0xFF;
        }

        if (control)
        {
            keyboardState[(int)VirtualKeyCode.CONTROL] = 0xFF;
        }

        if (shift)
        {
            keyboardState[(int)VirtualKeyCode.MENU] = 0xFF;
        }

        _ = User32.ToUnicode((uint)key, GetScanCode(key), keyboardState, buffer, 256, 0);

        return buffer.ToString();
    }
}
