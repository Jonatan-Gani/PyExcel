#if NETFRAMEWORK
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PyExcel.Forms;

/// <summary>
/// Common base for every PyExcel WinForms dialog. The dialogs are laid out in
/// code at absolute 96-DPI pixel coordinates; on a higher-DPI / smaller laptop
/// screen the system font renders larger but those fixed coordinates don't, so
/// rows collide and text overlaps — the "everything's crammed" report.
///
/// <para>WinForms' own auto-scaling (<see cref="AutoScaleMode.Font"/> /
/// <see cref="AutoScaleMode.Dpi"/>) is unreliable here: hosted inside an
/// Excel <c>.xll</c> without an app-level DPI manifest, its DPI/font
/// measurement can read 96 even on a 150% display, so it scales by 1.0 and
/// does nothing. Both earlier attempts hit exactly that. So this base does the
/// scaling itself, from the real monitor DPI reported by Win32
/// <c>GetDpiForWindow</c> (which honours the host's DPI awareness and can't get
/// stuck at 96), scaling the whole control tree up to match the font and then
/// capping the window to the screen so it always fits.</para>
/// </summary>
public abstract class ScaledForm : Form
{
    private bool _scaled;

    /// <summary>The DPI scale factor applied to this form (1.0 at 96 DPI,
    /// 1.5 at 150%, …). Available after the form has loaded; derived forms
    /// that add controls dynamically can use <see cref="ScaleNewControl"/> to
    /// bring them up to the same scale.</summary>
    protected float DpiScaleFactor { get; private set; } = 1f;

    protected ScaledForm()
    {
        // We scale manually in OnLoad, so turn WinForms' own (unreliable here)
        // auto-scaling off to avoid double-scaling or a fight over the layout.
        AutoScaleMode = AutoScaleMode.None;
        Font = SystemFonts.MessageBoxFont;
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        ApplyDpiScaling();
        FitToScreen();
    }

    private void ApplyDpiScaling()
    {
        if (_scaled) return;
        _scaled = true;

        var factor = GetDpiFactor();
        DpiScaleFactor = factor;
        if (factor <= 1.01f) return; // 100% — or a DPI-unaware host, where the OS bitmap-scales for us.

        SuspendLayout();
        ScaleControlTree(this, factor);
        ClientSize = Scale(ClientSize, factor);
        if (MinimumSize != Size.Empty) MinimumSize = Scale(MinimumSize, factor);
        if (MaximumSize != Size.Empty) MaximumSize = Scale(MaximumSize, factor);
        ResumeLayout(true);
    }

    /// <summary>Scale a control (and its descendants) that a derived form added
    /// after load, so dynamically-created rows match the form's DPI scale.
    /// No-op at 100%.</summary>
    protected void ScaleNewControl(Control control)
    {
        if (control is null || DpiScaleFactor <= 1.01f) return;
        var b = control.Bounds;
        control.Bounds = new Rectangle(
            Round(b.X * DpiScaleFactor), Round(b.Y * DpiScaleFactor),
            Round(b.Width * DpiScaleFactor), Round(b.Height * DpiScaleFactor));
        ScaleControlTree(control, DpiScaleFactor);
    }

    /// <summary>Recursively scale every child's bounds by <paramref name="factor"/>.
    /// Fonts are deliberately left alone: point sizes already render at the
    /// device DPI, so only the layout (positions and box sizes) needs to grow
    /// to match. AutoSize controls ignore the size part and re-fit to the font;
    /// their scaled position is what keeps them from colliding.</summary>
    private static void ScaleControlTree(Control parent, float factor)
    {
        foreach (Control c in parent.Controls)
        {
            var b = c.Bounds;
            c.Bounds = new Rectangle(
                Round(b.X * factor), Round(b.Y * factor),
                Round(b.Width * factor), Round(b.Height * factor));
            ScaleControlTree(c, factor);
        }
    }

    /// <summary>Keep the (now-scaled) window inside the screen's working area so
    /// it never opens bigger than the laptop screen; turn on scrolling so any
    /// overflow is still reachable.</summary>
    private void FitToScreen()
    {
        var area = Screen.FromControl(this).WorkingArea;
        var w = Math.Min(Width, area.Width);
        var h = Math.Min(Height, area.Height);
        if (w == Width && h == Height) return;
        AutoScroll = true;
        Size = new Size(w, h);
    }

    private float GetDpiFactor()
    {
        try
        {
            if (IsHandleCreated)
            {
                var dpi = GetDpiForWindow(Handle);
                if (dpi >= 48 && dpi <= 1200) return dpi / 96f;
            }
        }
        catch { /* GetDpiForWindow missing pre-Win10 1607 — fall back below */ }

        try
        {
            using var g = CreateGraphics();
            return g.DpiX / 96f;
        }
        catch { return 1f; }
    }

    private static Size Scale(Size s, float factor) => new(Round(s.Width * factor), Round(s.Height * factor));

    private static int Round(float v) => (int)Math.Round(v);

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);
}
#endif
