using System;
using System.Runtime.InteropServices;
public static class MeetingWindow {
    [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);

    public static bool AcknowledgeButton(IntPtr dialog, IntPtr button) {
        return button != IntPtr.Zero && GetParent(button) == dialog && IsWindowEnabled(button)
            && PostMessage(button, 0x00F5 /* BM_CLICK */, IntPtr.Zero, IntPtr.Zero);
    }

    [ComImport, Guid("00000114-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleWindow {
        void GetWindow(out IntPtr hwnd);
        void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool mode);
    }
    public static long Handle(object inspector) {
        IntPtr hwnd;
        ((IOleWindow)inspector).GetWindow(out hwnd);
        if (hwnd == IntPtr.Zero) throw new InvalidOperationException("Meeting window is not available.");
        return hwnd.ToInt64();
    }
}
