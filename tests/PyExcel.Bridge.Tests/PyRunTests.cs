using System;
using System.Collections.Generic;
using System.IO;
using PyExcel.Bridge;
using PyExcel.Excel;
using PyExcel.Kernel.Client;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// End-to-end tests for <see cref="PyRun.Execute"/> against a real Python
/// kernel. These are the cross-language conformance check for
/// <see cref="ArrowMarshal"/> ↔ <c>arrow_io.py</c>: every test in here
/// crosses the wire in both directions, so any disagreement on shape
/// metadata, type inference, or null handling surfaces immediately.
///
/// <para>Pairs with the in-process unit tests in
/// <see cref="ArrowMarshalTests"/>, which exercise the C# side in
/// isolation. Together they cover the matrix.</para>
/// </summary>
public class PyRunTests
{
    // -------------------------------------------------------------------------
    // Input encoding classification — pure-logic, no kernel needed.
    // -------------------------------------------------------------------------

    [Fact]
    public void EncodeInput_NullInput_YieldsNoArgs()
    {
        Assert.Null(PyRun.EncodeInput(null));
    }

    [Fact]
    public void EncodeInput_TwoDimensionalArray_EncodesAsTable()
    {
        var buf = PyRun.EncodeInput(new object?[,] { { 1.0, 2.0 }, { 3.0, 4.0 } })!;
        Assert.Equal(ArrowShape.Table, ArrowMarshal.PeekShape(buf).Shape);
    }

    [Fact]
    public void EncodeInput_OneDimensionalArray_EncodesAsVector()
    {
        var buf = PyRun.EncodeInput(new object?[] { 1.0, 2.0, 3.0 })!;
        Assert.Equal(ArrowShape.Vector, ArrowMarshal.PeekShape(buf).Shape);
    }

    [Fact]
    public void EncodeInput_Scalar_EncodesAsScalar()
    {
        var buf = PyRun.EncodeInput(42.0)!;
        Assert.Equal(ArrowShape.Scalar, ArrowMarshal.PeekShape(buf).Shape);
    }

    // -------------------------------------------------------------------------
    // End-to-end through a real kernel
    // -------------------------------------------------------------------------

    [Fact]
    public void Execute_NoInput_ReturnsScalarFromUserFunction()
    {
        using var fx = new KernelFixture();
        var script = fx.WriteScript("answer.py",
            "def transform():\n    return 42\n");

        var result = PyRun.Execute(
            script: script,
            input: null,
            kwargs: null,
            client: fx.Client);

        Assert.Equal(42.0, result);
    }

    [Fact]
    public void Execute_ScalarInput_RoundTripsThroughKernel()
    {
        using var fx = new KernelFixture();
        var script = fx.WriteScript("addone.py",
            "def transform(x):\n    return x + 1\n");

        var result = PyRun.Execute(
            script: script,
            input: 41.0,
            kwargs: null,
            client: fx.Client);

        Assert.Equal(42.0, result);
    }

    [Fact]
    public void Execute_StringScalar_RoundTrips()
    {
        using var fx = new KernelFixture();
        var script = fx.WriteScript("upper.py",
            "def transform(s):\n    return s.upper()\n");

        var result = PyRun.Execute(
            script: script,
            input: "hello",
            kwargs: null,
            client: fx.Client);

        Assert.Equal("HELLO", result);
    }

    [Fact]
    public void Execute_VectorInput_RoundTripsAsColumnSpill()
    {
        using var fx = new KernelFixture();
        var script = fx.WriteScript("rev.py",
            "def transform(xs):\n    return list(reversed(xs))\n");

        var result = PyRun.Execute(
            script: script,
            input: new object?[] { 1.0, 2.0, 3.0 },
            kwargs: null,
            client: fx.Client);

        // Default Python vector orientation is column → N×1 rectangle.
        var rect = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, rect.GetLength(0));
        Assert.Equal(1, rect.GetLength(1));
        Assert.Equal(3.0, rect[0, 0]);
        Assert.Equal(2.0, rect[1, 0]);
        Assert.Equal(1.0, rect[2, 0]);
    }

    [Fact]
    public void Execute_TableInput_RoundTripsAsTable()
    {
        using var fx = new KernelFixture();
        var script = fx.WriteScript("double.py",
            "def transform(df):\n    return df * 2\n");

        var result = PyRun.Execute(
            script: script,
            input: new object?[,]
            {
                { 1.0, 2.0 },
                { 3.0, 4.0 },
            },
            kwargs: null,
            client: fx.Client);

        var table = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, table.GetLength(0));
        Assert.Equal(2, table.GetLength(1));
        Assert.Equal(2.0, table[0, 0]);
        Assert.Equal(4.0, table[0, 1]);
        Assert.Equal(6.0, table[1, 0]);
        Assert.Equal(8.0, table[1, 1]);
    }

    [Fact]
    public void Execute_KwargsArePassedThrough()
    {
        using var fx = new KernelFixture();
        var script = fx.WriteScript("kw.py",
            "def transform(x, *, factor):\n    return x * factor\n");

        var result = PyRun.Execute(
            script: script,
            input: 8.0,
            kwargs: new Dictionary<string, object?> { ["factor"] = 5L },
            client: fx.Client);

        Assert.Equal(40.0, result);
    }

    [Fact]
    public void Execute_NoneReturn_YieldsEmptyResultSentinel()
    {
        using var fx = new KernelFixture();
        var script = fx.WriteScript("none.py",
            "def transform():\n    return None\n");

        var result = PyRun.Execute(
            script: script,
            input: null,
            kwargs: null,
            client: fx.Client);

        Assert.Same(PyRun.EmptyResult, result);
    }

    [Fact]
    public void Execute_UserException_BubblesAsKernelException()
    {
        using var fx = new KernelFixture();
        var script = fx.WriteScript("boom.py",
            "def transform():\n    raise ValueError('boom')\n");

        var ex = Assert.Throws<KernelException>(() =>
            PyRun.Execute(
                script: script,
                input: null,
                kwargs: null,
                client: fx.Client));

        Assert.Equal("Exception", ex.Code);
        Assert.Equal("ValueError", ex.PythonType);
    }

    [Fact]
    public void Execute_RelativeScript_ResolvedAgainstWorkbookDirectory()
    {
        using var fx = new KernelFixture();
        // Write the script under the fixture dir but pass a *relative* path
        // and the fixture dir as workbookDirectory.
        fx.WriteScript("rel.py",
            "def transform():\n    return 7\n");

        var result = PyRun.Execute(
            script: "rel.py",
            input: null,
            kwargs: null,
            client: fx.Client,
            workbookDirectory: fx.ScratchDir);

        Assert.Equal(7.0, result);
    }

    [Fact]
    public void Execute_CustomFunction_HonoredOverDefault()
    {
        using var fx = new KernelFixture();
        var script = fx.WriteScript("named.py",
            "def transform():\n    return 'wrong'\n" +
            "def go():\n    return 'right'\n");

        var result = PyRun.Execute(
            script: script,
            input: null,
            kwargs: null,
            client: fx.Client,
            function: "go");

        Assert.Equal("right", result);
    }

    // -------------------------------------------------------------------------
    // Multi-input dispatch (ExecuteMany) — the ribbon-button path
    // -------------------------------------------------------------------------

    [Fact]
    public void ExecuteMany_TwoScalarArgs_PassedPositionally()
    {
        using var fx = new KernelFixture();
        var script = fx.WriteScript("add2.py",
            "def transform(a, b):\n    return a + b\n");

        var result = PyRun.ExecuteMany(
            script: script,
            inputs: new object?[] { 40.0, 2.0 },
            kwargs: null,
            client: fx.Client);

        Assert.Equal(42.0, result);
    }

    [Fact]
    public void ExecuteMany_MixedShapes_PreservesOrder()
    {
        using var fx = new KernelFixture();
        // First arg is a vector, second a scalar; the function indexes the
        // vector and multiplies by the scalar — so a misordering would give
        // a different answer.
        var script = fx.WriteScript("vecscale.py",
            "def transform(xs, factor):\n    return [x * factor for x in xs]\n");

        var result = PyRun.ExecuteMany(
            script: script,
            inputs: new object?[] { new object?[] { 1.0, 2.0, 3.0 }, 10.0 },
            kwargs: null,
            client: fx.Client);

        var rect = Assert.IsType<object?[,]>(result);
        Assert.Equal(10.0, rect[0, 0]);
        Assert.Equal(20.0, rect[1, 0]);
        Assert.Equal(30.0, rect[2, 0]);
    }

    [Fact]
    public void ExecuteMany_NoInputs_CallsNoArgFunction()
    {
        using var fx = new KernelFixture();
        var script = fx.WriteScript("noarg.py",
            "def transform():\n    return 5\n");

        var result = PyRun.ExecuteMany(
            script: script,
            inputs: Array.Empty<object?>(),
            kwargs: null,
            client: fx.Client);

        Assert.Equal(5.0, result);
    }

    [Fact]
    public void ExecuteMany_NullInputInList_Throws()
    {
        using var fx = new KernelFixture();
        var ex = Assert.Throws<ArgumentException>(() =>
            PyRun.ExecuteMany(
                script: "x.py",
                inputs: new object?[] { 1.0, null, 3.0 },
                kwargs: null,
                client: fx.Client));
        Assert.Contains("index 1", ex.Message);
    }

    [Fact]
    public void ExecuteMany_NullInputsList_Throws()
    {
        using var fx = new KernelFixture();
        Assert.Throws<ArgumentNullException>(() =>
            PyRun.ExecuteMany("x.py", null!, null, fx.Client));
    }

    // -------------------------------------------------------------------------
    // Argument validation (no kernel needed)
    // -------------------------------------------------------------------------

    [Fact]
    public void Execute_NullScript_Throws()
    {
        using var fx = new KernelFixture();
        Assert.Throws<ArgumentNullException>(() =>
            PyRun.Execute(null!, null, null, fx.Client));
    }

    [Fact]
    public void Execute_EmptyScript_Throws()
    {
        using var fx = new KernelFixture();
        Assert.Throws<ArgumentException>(() =>
            PyRun.Execute("", null, null, fx.Client));
    }

    [Fact]
    public void Execute_NullClient_Throws()
    {
        Assert.Throws<ArgumentNullException>(() =>
            PyRun.Execute("x.py", null, null, null!));
    }

    // -------------------------------------------------------------------------
    // Fixture — spawns one kernel + scratch dir, cleans both up.
    // -------------------------------------------------------------------------

    private sealed class KernelFixture : IDisposable
    {
        public KernelSupervisor Supervisor { get; }
        public KernelClient Client { get; }
        public string ScratchDir { get; }

        public KernelFixture()
        {
            // Use PythonResolver so the test exercises the same discovery
            // path the .xll will use in production.
            var python = PythonResolver.ResolvePython();
            var embedded = PythonResolver.ResolveEmbeddedPath();

            ScratchDir = Path.Combine(
                Path.GetTempPath(),
                "pyexcel-pyruntest-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(ScratchDir);

            Supervisor = KernelSupervisor.StartPython(python, embedded);
            Client = new KernelClient(Supervisor);
        }

        public string WriteScript(string filename, string body)
        {
            var path = Path.Combine(ScratchDir, filename);
            File.WriteAllText(path, body);
            return path;
        }

        public void Dispose()
        {
            try { Supervisor.Dispose(); } catch { /* best-effort */ }
            try { Directory.Delete(ScratchDir, recursive: true); } catch { /* best-effort */ }
        }
    }
}
