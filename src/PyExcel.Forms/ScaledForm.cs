#if NETFRAMEWORK
using System.Drawing;
using System.Windows.Forms;

namespace PyExcel.Forms;

/// <summary>
/// Common base for every PyExcel WinForms dialog. Makes a dialog laid out in
/// code render at a consistent size on every machine — scaling the whole
/// layout up on a high-DPI laptop instead of cramming the fixed-pixel layout
/// under an enlarged system font.
///
/// <para>WinForms' default is <see cref="AutoScaleMode.Font"/> but with no
/// design-time baseline (<c>AutoScaleDimensions</c> defaults to empty), so the
/// scale factor is always 1 and nothing scales — the original "crammed on my
/// laptop" report. An earlier attempt used <see cref="AutoScaleMode.Dpi"/>,
/// but that keys off <see cref="Control.DeviceDpi"/>, which stays 96 unless
/// the host opts into WinForms per-monitor DPI (Excel + .NET Framework 4.8
/// does not), so it was a no-op while the font — measured from the real
/// device context — still grew, leaving the layout crammed.</para>
///
/// <para>This pins <see cref="AutoScaleMode.Font"/> with the standard Segoe UI
/// 9 pt / 96-DPI baseline (7×15): Font mode measures the form's font against
/// the real device context, so it scales the whole control tree to match the
/// runtime font size regardless of the host's DPI-awareness opt-in. Every
/// hardcoded <c>ClientSize</c> / control coordinate is then a 96-DPI logical
/// value that grows proportionally on a high-DPI screen.</para>
/// </summary>
public abstract class ScaledForm : Form
{
    protected ScaledForm()
    {
        // Order matters and mirrors what the WinForms designer emits: establish
        // the font the layout is measured against, then the matching 96-DPI
        // baseline, then the mode — all before any derived control is added, so
        // the first layout scales the whole tree to the runtime font size.
        Font = SystemFonts.MessageBoxFont;
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
    }
}
#endif
