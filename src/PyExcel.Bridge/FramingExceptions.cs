using System;

namespace PyExcel.Bridge;

/// <summary>Base for all framing-layer failures.</summary>
public class FramingException : Exception
{
    public FramingException(string message) : base(message) { }
    public FramingException(string message, Exception inner) : base(message, inner) { }
}

/// <summary>Frame size exceeds the configured cap (caught before allocating).</summary>
public sealed class FrameTooLargeException : FramingException
{
    public FrameTooLargeException(string message) : base(message) { }
}

/// <summary>Underlying stream closed before the full frame was read.</summary>
public sealed class TruncatedFrameException : FramingException
{
    public TruncatedFrameException(string message) : base(message) { }
    public TruncatedFrameException(string message, Exception inner) : base(message, inner) { }
}

/// <summary>Frame header parsed but internal lengths are inconsistent.</summary>
public sealed class MalformedFrameException : FramingException
{
    public MalformedFrameException(string message) : base(message) { }
    public MalformedFrameException(string message, Exception inner) : base(message, inner) { }
}
