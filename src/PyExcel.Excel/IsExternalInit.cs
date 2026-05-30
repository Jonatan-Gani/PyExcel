// C# 9+ records and init-only setters compile down to references to
// System.Runtime.CompilerServices.IsExternalInit. .NET 5+ ships it in
// the BCL; netstandard2.0 and net48 don't. The C# compiler lets any
// assembly declare it instead, so we provide an internal polyfill here.
//
// Mirrors the same polyfill in PyExcel.State — needed independently per
// assembly because the type is consumed at compile time and the runtime
// only looks it up via the declaring assembly.

#if !NET5_0_OR_GREATER
namespace System.Runtime.CompilerServices
{
    using System.ComponentModel;

    [EditorBrowsable(EditorBrowsableState.Never)]
    internal static class IsExternalInit { }
}
#endif
