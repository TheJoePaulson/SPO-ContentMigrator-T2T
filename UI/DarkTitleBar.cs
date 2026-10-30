using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace ContentMigrator.UI;

/// <summary>
/// Turns the Windows title bar dark to match the theme. Windows 10 (20H1+) gets the dark
/// title bar; Windows 11 also gets the exact caption and text colours. Any failure is
/// ignored, leaving the normal title bar, so the app always starts.
/// </summary>
internal static class DarkTitleBar
{
    private const int DwmwaUseImmersiveDarkModeBefore20H1 = 19;
    private const int DwmwaUseImmersiveDarkMode = 20;
    private const int DwmwaCaptionColor = 35;
    private const int DwmwaTextColor = 36;

    // COLORREF is 0x00BBGGRR.
    private const int CaptionColor = 0x00332D2A; // #2A2D33 (Brush.Surface)
    private const int TextColor = 0x00EBE8E6;    // #E6E8EB (Brush.Text)

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

    public static void Attach(Window window)
    {
        window.SourceInitialized += (_, _) => Apply(window);
    }

    private static void Apply(Window window)
    {
        try
        {
            var hwnd = new WindowInteropHelper(window).Handle;
            if (hwnd == IntPtr.Zero) return;

            var on = 1;
            if (DwmSetWindowAttribute(hwnd, DwmwaUseImmersiveDarkMode, ref on, sizeof(int)) != 0)
            {
                DwmSetWindowAttribute(hwnd, DwmwaUseImmersiveDarkModeBefore20H1, ref on, sizeof(int));
            }

            var caption = CaptionColor;
            DwmSetWindowAttribute(hwnd, DwmwaCaptionColor, ref caption, sizeof(int));
            var text = TextColor;
            DwmSetWindowAttribute(hwnd, DwmwaTextColor, ref text, sizeof(int));
        }
        catch
        {
            // Cosmetic only.
        }
    }
}
