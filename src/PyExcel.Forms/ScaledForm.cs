#if NETFRAMEWORK
using System.Drawing;
using System.Windows.Forms;

namespace PyExcel.Forms;

/// <summary>
/// Common base for every PyExcel WinForms dialog. Pins DPI auto-scaling to a
/// fixed 96-DPI baseline so a dialog laid out in code renders at the same
/// <em>logical</em> size on every machine — scaling up cleanly on a high-DPI
/// laptop instead of cramming the fixed-pixel layout under an enlarged system
/// font.
///
/// <para>Without this, the forms inherit WinForms' default
/// <see cref="AutoScaleMode.Font"/> with no design-time baseline. On a
/// high-DPI display the font then grows but the hand-coded control bounds
/// don't, so labels clip and buttons overlap — the "it's crammed on my
/// laptop, and different on every screen" report. <see cref="AutoScaleMode.Dpi"/>
/// against a fixed 96-DPI baseline scales the whole layout by the monitor's
/// DPI, which is exactly the per-screen difference being seen, so every
/// hardcoded <c>ClientSize</c> / control coordinate is interpreted as a
/// 96-DPI value and blown up to match the runtime DPI.</para>
/// </summary>
public abstract class ScaledForm : Form
{
    protected ScaledForm()
    {
        // Establish the 96-DPI design baseline before any controls are added,
        // then let WinForms scale the layout to the runtime monitor DPI.
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
    }
}
#endif
