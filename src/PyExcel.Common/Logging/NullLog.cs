using System;

namespace PyExcel.Common.Logging;

/// <summary>No-op logger for tests and headless scenarios.</summary>
public sealed class NullLog : ILog
{
    public static readonly NullLog Instance = new();
    public void Trace(string message) { }
    public void Debug(string message) { }
    public void Info(string message) { }
    public void Warn(string message) { }
    public void Error(string message, Exception? exception = null) { }
}
