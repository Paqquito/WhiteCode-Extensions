using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

public static class ZoomDialog {
    [StructLayout(LayoutKind.Sequential)] struct Rect { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll")] static extern bool IsWindowEnabled(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")] static extern IntPtr GetWindowDpiAwarenessContext(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr hwnd, IntPtr dc, uint flags);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr w, IntPtr l);

    public const int FooterScale = 2;

    public static MemoryStream CaptureFooter(IntPtr hwnd, out int width, out int height, out int top) {
        var previous = SetThreadDpiAwarenessContext(GetWindowDpiAwarenessContext(hwnd));
        try {
            Rect r;
            if (!GetClientRect(hwnd, out r) || r.Right < 100 || r.Bottom < 100 || r.Right > 4096 || r.Bottom > 4096)
                throw new InvalidOperationException("Unexpected Zoom dialog size.");
            width = r.Right;
            height = r.Bottom;
            top = height * 3 / 4;
            using (var bitmap = new Bitmap(width, height))
            using (var graphics = Graphics.FromImage(bitmap)) {
                var dc = graphics.GetHdc();
                try { if (!PrintWindow(hwnd, dc, 3)) throw new InvalidOperationException("Zoom dialog capture failed."); }
                finally { graphics.ReleaseHdc(dc); }
                var footer = new Rectangle(0, top, width, height - top);
                using (var enlarged = new Bitmap(width * FooterScale, footer.Height * FooterScale))
                using (var drawing = Graphics.FromImage(enlarged)) {
                    drawing.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    drawing.DrawImage(bitmap, new Rectangle(0, 0, enlarged.Width, enlarged.Height), footer, GraphicsUnit.Pixel);
                    var stream = new MemoryStream();
                    enlarged.Save(stream, ImageFormat.Png);
                    stream.Position = 0;
                    return stream;
                }
            }
        } finally { SetThreadDpiAwarenessContext(previous); }
    }

    public static bool ClickClient(IntPtr hwnd, int width, int height, int x, int y) {
        var previous = SetThreadDpiAwarenessContext(GetWindowDpiAwarenessContext(hwnd));
        try {
            Rect r;
            if (!IsWindowVisible(hwnd) || !IsWindowEnabled(hwnd) || !GetClientRect(hwnd, out r) ||
                width != r.Right || height != r.Bottom || x < 0 || y < 0 || x >= width || y >= height) return false;
            var packed = new IntPtr((y << 16) | (x & 65535));
            return PostMessage(hwnd, 0x200, IntPtr.Zero, packed) && PostMessage(hwnd, 0x201, new IntPtr(1), packed) && PostMessage(hwnd, 0x202, IntPtr.Zero, packed);
        } finally { SetThreadDpiAwarenessContext(previous); }
    }
}
