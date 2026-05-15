using System;

namespace PyExcel.Common.Logging;

/// <summary>
/// Minimal logging surface used across all PyExcel projects.
///
/// Kept deliberately tiny — five severity methods, no scopes, no structured
/// fields beyond <c>string.Format</c>-style messages. We will not adopt
/// Microsoft.Extensions.Logging here: a .xll add-in cannot afford the
/// dependency surface, and PyExcel logs always end up in the same
/// <c>%TEMP%\PyExcel_Debug.log</c> file regardless of severity. The richer
/// shape can be added in v2.1 if it ever justifies its cost.
/// </summary>
public interface ILog
{
    void Trace(string message);
    void Debug(string message);
    void Info(string message);
    void Warn(string message);
    void Error(string message, Exception? exception = null);
}
