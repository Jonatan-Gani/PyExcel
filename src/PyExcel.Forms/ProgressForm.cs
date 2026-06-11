#if NETFRAMEWORK
using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using PyExcel.Excel;

namespace PyExcel.Forms;

/// <summary>
/// The non-blocking run-progress dialog (Phase 8 — closes the long-deferred
/// Phase 4 #5). Shown modeless while the kernel runs on a background
/// thread; it renders the <c>PROGRESS</c> frames <see cref="RangeRunner"/>
/// forwards and offers a working Cancel that trips a
/// <see cref="CancellationToken"/> threaded into the async run (which the
/// kernel turns into a <c>CANCEL</c> frame).
///
/// <para>Implements the cross-platform <see cref="IRunProgressSink"/> so
/// <see cref="RangeRunner"/> drives it without a WinForms dependency. All
/// mutators marshal onto the UI thread — the kernel's progress events
/// arrive on a background thread.</para>
/// </summary>
public sealed class ProgressForm : Form, IRunProgressSink
{
    private readonly ProgressBar _bar;
    private readonly Label _status;
    private readonly Button _cancelButton;
    private readonly CancellationTokenSource _cts = new();
    private bool _completedNormally;

    public CancellationToken CancellationToken => _cts.Token;

    /// <summary>Create and show the dialog modeless, owned by Excel's
    /// window, and return it as the run's progress sink.</summary>
    public static ProgressForm StartModeless(IWin32Window? owner, string title)
    {
        var form = new ProgressForm(title);
        if (owner is null) form.Show(); else form.Show(owner);
        return form;
    }

    private ProgressForm(string title)
    {
        Text = title;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ControlBox = false; // force Cancel rather than the X
        Font = SystemFonts.MessageBoxFont;
        ClientSize = new Size(360, 96);

        _status = new Label
        {
            Left = 12,
            Top = 12,
            Width = ClientSize.Width - 24,
            Height = 16,
            Text = "Working…",
        };
        Controls.Add(_status);

        _bar = new ProgressBar
        {
            Left = 12,
            Top = 34,
            Width = ClientSize.Width - 24,
            Height = 18,
            Style = ProgressBarStyle.Marquee,
            MarqueeAnimationSpeed = 30,
        };
        Controls.Add(_bar);

        _cancelButton = new Button
        {
            Text = "Cancel",
            Left = ClientSize.Width - 92,
            Top = ClientSize.Height - 32,
            Width = 80,
        };
        _cancelButton.Click += (_, _) => RequestCancel();
        Controls.Add(_cancelButton);
        CancelButton = _cancelButton;
    }

    public void Report(double? percent, string? message)
    {
        if (IsDisposed) return;
        if (InvokeRequired)
        {
            try { BeginInvoke(new Action(() => Report(percent, message))); } catch { }
            return;
        }

        _status.Text = ProgressModel.FormatLine(percent, message);
        if (percent is null)
        {
            _bar.Style = ProgressBarStyle.Marquee;
        }
        else
        {
            _bar.Style = ProgressBarStyle.Continuous;
            _bar.Value = ProgressModel.ClampPercent(percent.Value);
        }
    }

    public void Complete()
    {
        if (IsDisposed) return;
        if (InvokeRequired)
        {
            try { BeginInvoke(new Action(Complete)); } catch { }
            return;
        }

        _completedNormally = true;
        Close();
    }

    private void RequestCancel()
    {
        if (!_cts.IsCancellationRequested) _cts.Cancel();
        _status.Text = "Cancelling…";
        _cancelButton.Enabled = false;
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        // A close that wasn't our own Complete() (e.g. Alt+F4) counts as a
        // cancel request so the run doesn't keep going headless.
        if (!_completedNormally && !_cts.IsCancellationRequested)
            _cts.Cancel();
        base.OnFormClosed(e);
    }
}
#endif
