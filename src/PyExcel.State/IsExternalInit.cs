// C# 9+ records and init-only setters compile down to references to
// System.Runtime.CompilerServices.IsExternalInit. .NET 5+ ships it in
// the BCL; netstandard2.0 and net48 don't. The C# compiler lets any
// assembly declare it instead, so we provide an internal polyfill here.
//
// Once the project drops net48 and netstandard2.0 in favour of a target
// with the type built in, this file can be deleted — no other source
// references it directly.

#if !NET5_0_OR_GREATER
namespace System.Runtime.CompilerServices
{
    using System.ComponentModel;

    [EditorBrowsable(EditorBrowsableState.Never)]
    internal static class IsExternalInit { }
}
#endif
