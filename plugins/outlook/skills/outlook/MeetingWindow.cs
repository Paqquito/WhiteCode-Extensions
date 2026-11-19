using System;
using System.Runtime.InteropServices;
public static class MeetingWindow {
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
