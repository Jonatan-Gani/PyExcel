#if NETFRAMEWORK
using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using PyExcel.Common.Logging;
using PyExcel.Setup;

namespace PyExcel.Forms;

/// <summary>
/// The first-run setup wizard (Phase 9 — the runtime entry point for the
/// Phase 7 headless engine). Drives <see cref="SetupService"/> end to end —
/// probe Python, create the per-project <c>.pyexcel-venv</c>, extract the
/// kernel, pip-install the requirements, verify the imports — and streams
/// every venv/pip line into a live log so the user (and a support log) can
/// see exactly what happened. This is the "surface venv/pip output" the
/// roadmap's Setup-diagnostics item asks for.
///
/// <para>The headless <see cref="SetupService"/> already runs on Linux CI and
/// is unit-tested; <see cref="SetupReport"/> formats its result and is tested
/// too. This form is the thin net48 WinForms shell that hosts them — the
/// Windows smoke test, same split as every other dialog.</para>
/// </summary>
public sealed class SetupForm : Form
{
    private readonly string _projectPath;
    private readonly ILog? _innerLog;
    private readonly TextBox _log;
    private readonly Label _status;
    private readonly ProgressBar _bar;
    private readonly Button _closeButton;
    private SetupResult? _result;

    /// <summary>Run the wizard modally against <paramref name="projectPath"/>
    /// (the workbook directory the venv is provisioned next to, matching
    /// <c>PyExcel.Excel.PythonResolver</c>'s layout). Returns the run's success
    /// flag once it finishes, or <see langword="null"/> if the dialog was
    /// dismissed before completion. Forwards every log line to
    /// <paramref name="log"/> too, so the run is also captured in
    /// <c>%TEMP%\PyExcel_Debug.log</c>. Returns <c>bool?</c> rather than the
    /// <see cref="SetupResult"/> so callers (the ribbon) don't take a
    /// <c>PyExcel.Setup</c> reference of their own.</summary>
    public static bool? Run(IWin32Window? owner, string projectPath, ILog? log = null)
    {
        using var form = new SetupForm(projectPath, log);
        if (owner is null) form.ShowDialog(); else form.ShowDialog(owner);
        return form._result?.Success;
    }

    private SetupForm(string projectPath, ILog? log)
    {
        _projectPath = projectPath ?? throw new ArgumentNullException(nameof(projectPath));
        _innerLog = log;

        Text = "PyExcel Setup";
        FormBorderStyle = FormBorderStyle.Sizable;
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ControlBox = false; // no X while running; enabled on completion
        Font = SystemFonts.MessageBoxFont;
        ClientSize = new Size(580, 420);
        MinimumSize = new Size(420, 280);

        _status = new Label
        {
            Left = 12,
            Top = 12,
            Width = ClientSize.Width - 24,
            Height = 18,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Text = "Setting up the Python environment…",
        };
        Controls.Add(_status);

        _bar = new ProgressBar
        {
            Left = 12,
            Top = 34,
            Width = ClientSize.Width - 24,
            Height = 16,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Style = ProgressBarStyle.Marquee,
            MarqueeAnimationSpeed = 30,
        };
        Controls.Add(_bar);

        _log = new TextBox
        {
            Left = 12,
            Top = 58,
            Width = ClientSize.Width - 24,
            Height = ClientSize.Height - 58 - 44,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            Font = new Font(FontFamily.GenericMonospace, 8.5f),
            BackColor = SystemColors.Window,
        };
        Controls.Add(_log);

        _closeButton = new Button
        {
            Text = "Close",
            Width = 88,
            Height = 28,
            Left = ClientSize.Width - 100,
            Top = ClientSize.Height - 36,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            Enabled = false, // until the run finishes
        };
        _closeButton.Click += (_, _) => Close();
        Controls.Add(_closeButton);
        AcceptButton = _closeButton;
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        // Kick the headless pipeline off the UI thread; stream its log here.
        var sink = new SinkLog(AppendLine, _innerLog);
        Task.Run(() => new SetupService(sink).Run(_projectPath))
            .ContinueWith(t => OnFinished(t), TaskScheduler.Default);
    }

    private void OnFinished(Task<SetupResult> t)
    {
        if (IsDisposed) return;
        if (InvokeRequired) { try { BeginInvoke(new Action(() => OnFinished(t))); } catch { } return; }

        _bar.Style = ProgressBarStyle.Continuous;
        _bar.Value = _bar.Maximum;

        if (t.IsFaulted)
        {
            _status.Text = "Setup failed unexpectedly.";
            AppendLine("ERROR: " + (t.Exception?.GetBaseException().Message ?? "unknown error"), true);
        }
        else
        {
            _result = t.Result;
            _status.Text = SetupReport.Headline(_result);
            AppendLine(string.Empty, false);
            AppendLine(SetupReport.Summarize(_result), false);
        }

        ControlBox = true;
        _closeButton.Enabled = true;
        CancelButton = _closeButton;
        _closeButton.Focus();
    }

    private void AppendLine(string line, bool isError)
    {
        if (IsDisposed) return;
        if (InvokeRequired) { try { BeginInvoke(new Action(() => AppendLine(line, isError))); } catch { } return; }
        _log.AppendText(line + Environment.NewLine);
    }

    /// <summary>An <see cref="ILog"/> that mirrors each line into the form's
    /// log box and forwards to an optional inner log (the file log) so the run
    /// lands in both places. Trace is dropped as too noisy for the UI.</summary>
    private sealed class SinkLog : ILog
    {
        private readonly Action<string, bool> _append;
        private readonly ILog? _inner;

        public SinkLog(Action<string, bool> append, ILog? inner)
        {
            _append = append;
            _inner = inner;
        }

        public void Trace(string message) => _inner?.Trace(message);
        public void Debug(string message) { _append(message, false); _inner?.Debug(message); }
        public void Info(string message) { _append(message, false); _inner?.Info(message); }
        public void Warn(string message) { _append("WARN: " + message, false); _inner?.Warn(message); }

        public void Error(string message, Exception? exception = null)
        {
            _append("ERROR: " + message + (exception is null ? string.Empty : " — " + exception.Message), true);
            _inner?.Error(message, exception);
        }
    }
}
#endif
