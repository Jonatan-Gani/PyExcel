using System;

namespace PyExcel.Bridge;

/// <summary>
/// One line the kernel subprocess wrote to its standard output or standard
/// error. Carried by <see cref="KernelSupervisor.OutputReceived"/> so the
/// host can surface user <c>print()</c> / <c>logging</c> output that would
/// otherwise vanish into the redirected-but-undrained child pipes.
/// </summary>
public sealed class KernelOutputEventArgs : EventArgs
{
    /// <summary>True when the line came from the child's standard error,
    /// false for standard output.</summary>
    public bool IsError { get; }

    /// <summary>The line text, with the trailing newline already stripped
    /// by <see cref="System.IO.TextReader.ReadLine"/>.</summary>
    public string Text { get; }

    public KernelOutputEventArgs(bool isError, string text)
    {
        IsError = isError;
        Text = text ?? string.Empty;
    }
}
