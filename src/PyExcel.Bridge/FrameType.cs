namespace PyExcel.Bridge;

/// <summary>
/// Stable frame-type identifiers shared with the Python kernel.
///
/// Numeric values are part of the on-wire protocol — new types append,
/// existing values never re-number. Must mirror <c>FrameType</c> in
/// <c>embedded/pyexcel/kernel/framing.py</c> byte-for-byte.
/// </summary>
public enum FrameType : byte
{
    Hello = 1,
    Ping = 2,
    Pong = 3,
    Shutdown = 4,
    Error = 5,

    RunRequest = 10,
    RunResult = 11,
    Progress = 12,
    Log = 13,
    Cancel = 14,
    ListJobs = 15,
}
