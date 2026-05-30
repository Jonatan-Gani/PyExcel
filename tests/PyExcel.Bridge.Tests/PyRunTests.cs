using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
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
    // Async overload — cancellation path is the new bit; happy path proves
    // the wrapper preserves the sync semantics.
    // -------------------------------------------------------------------------

    [Fact]
    public async Task ExecuteAsync_HappyPath_MatchesSyncResult()
    {
        using var fx = new KernelFixture();
        var script = fx.WriteScript("addone_async.py",
            "def transform(x):\n    return x + 1\n");

        var result = await PyRun.ExecuteAsync(
            script: script,
            input: 41.0,
            kwargs: null,
            client: fx.Client);

        Assert.Equal(42.0, result);
    }

    [Fact]
    public async Task ExecuteAsync_TokenCancelledMidRun_ThrowsOperationCanceled()
    {
        // Cooperative loop — polls is_cancelled() so a CANCEL frame from
        // the host can actually interrupt it. Exercises the UDF cancel
        // bridge end-to-end: token → KernelClient.Cancel → kernel CANCEL
        // frame → user script returns early → kernel ERROR/Cancelled →
        // host throws OperationCanceledException.
        using var fx = new KernelFixture();
        var script = fx.WriteScript("loop_async.py",
            "import time\n" +
            "from pyexcel.kernel import is_cancelled\n" +
            "def transform():\n" +
            "    for _ in range(200):\n" +
            "        if is_cancelled():\n" +
            "            return 'stopped'\n" +
            "        time.sleep(0.05)\n" +
            "    return 'finished'\n");

        using var cts = new CancellationTokenSource();
        var run = PyRun.ExecuteAsync(
            script: script,
            input: null,
            kwargs: null,
            client: fx.Client,
            cancellationToken: cts.Token);

        await Task.Delay(200);
        cts.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => run);
    }

    [Fact]
    public async Task ExecuteManyAsync_HappyPath_MatchesSyncResult()
    {
        using var fx = new KernelFixture();
        var script = fx.WriteScript("add2_async.py",
            "def transform(a, b):\n    return a + b\n");

        var result = await PyRun.ExecuteManyAsync(
            script: script,
            inputs: new object?[] { 40.0, 2.0 },
            kwargs: null,
            client: fx.Client);

        Assert.Equal(42.0, result);
    }

    // -------------------------------------------------------------------------
    // Cross-language date round-trip — the most subtle bit of the data
    // plane. Excel cells with a date format come through Value2 as doubles
    // (OADate), so they exercise the numeric path; the reverse direction
    // — a Python script returning a datetime / date — is what these tests
    // pin. Before the date/timestamp decoders landed, a script returning a
    // datetime would surface as the literal string "TimestampArray" in
    // Excel; these tests guard against that regression.
    // -------------------------------------------------------------------------

    [Fact]
    public void Execute_PythonReturnsDatetime_DecodesAsDateTime()
    {
        // pa.array([dt.datetime(...)]) → timestamp[us] with no timezone;
        // the C# decoder should land on a naive DateTime matching the
        // Python wall-clock value.
        using var fx = new KernelFixture();
        var script = fx.WriteScript("datetime_scalar.py",
            "import datetime as dt\n" +
            "def transform():\n" +
            "    return dt.datetime(2024, 1, 15, 9, 30, 45)\n");

        var result = PyRun.Execute(
            script: script,
            input: null,
            kwargs: null,
            client: fx.Client);

        var dt = Assert.IsType<DateTime>(result);
        Assert.Equal(new DateTime(2024, 1, 15, 9, 30, 45, DateTimeKind.Unspecified), dt);
        Assert.Equal(DateTimeKind.Unspecified, dt.Kind);
    }

    [Fact]
    public void Execute_PythonReturnsDate_DecodesAsMidnightDateTime()
    {
        // pa.array([dt.date(...)]) → date32 (days since epoch). Decoder
        // returns a midnight DateTime.
        using var fx = new KernelFixture();
        var script = fx.WriteScript("date_scalar.py",
            "import datetime as dt\n" +
            "def transform():\n" +
            "    return dt.date(2024, 1, 15)\n");

        var result = PyRun.Execute(
            script: script,
            input: null,
            kwargs: null,
            client: fx.Client);

        var dt = Assert.IsType<DateTime>(result);
        Assert.Equal(new DateTime(2024, 1, 15, 0, 0, 0, DateTimeKind.Unspecified), dt);
    }

    [Fact]
    public void Execute_PythonReturnsDatetimeList_DecodesAsDateTimeVector()
    {
        // A list of datetimes — the 1-D vector path. Exercises both the
        // timestamp decoder and the vector spill geometry.
        using var fx = new KernelFixture();
        var script = fx.WriteScript("datetime_list.py",
            "import datetime as dt\n" +
            "def transform():\n" +
            "    return [\n" +
            "        dt.datetime(2024, 1, 1, 0, 0, 0),\n" +
            "        dt.datetime(2024, 6, 15, 12, 30, 0),\n" +
            "        dt.datetime(2024, 12, 31, 23, 59, 59),\n" +
            "    ]\n");

        var result = PyRun.Execute(
            script: script,
            input: null,
            kwargs: null,
            client: fx.Client);

        // Column vector → N×1 rectangle.
        var rect = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, rect.GetLength(0));
        Assert.Equal(1, rect.GetLength(1));
        Assert.Equal(new DateTime(2024, 1, 1), (DateTime)rect[0, 0]!);
        Assert.Equal(new DateTime(2024, 6, 15, 12, 30, 0), (DateTime)rect[1, 0]!);
        Assert.Equal(new DateTime(2024, 12, 31, 23, 59, 59), (DateTime)rect[2, 0]!);
    }

    [Fact]
    public void Execute_PythonReturnsPandasTimestampSeries_DecodesAsDateTimeVector()
    {
        // A pandas Series of timestamps lands as timestamp[ns] (nanosecond
        // precision — pandas's default). The nanosecond arm of the
        // timestamp decoder is what we're exercising here. Both input
        // strings carry an explicit time component because pandas 2.x's
        // default to_datetime parser rejects mixed-format input
        // ('2024-01-01' + '2024-06-15 12:30:00' would need format='mixed').
        using var fx = new KernelFixture();
        var script = fx.WriteScript("pandas_ts.py",
            "import pandas as pd\n" +
            "def transform():\n" +
            "    return pd.Series(pd.to_datetime([\n" +
            "        '2024-01-01 00:00:00',\n" +
            "        '2024-06-15 12:30:00',\n" +
            "    ]))\n");

        var result = PyRun.Execute(
            script: script,
            input: null,
            kwargs: null,
            client: fx.Client);

        var rect = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, rect.GetLength(0));
        Assert.Equal(new DateTime(2024, 1, 1), (DateTime)rect[0, 0]!);
        Assert.Equal(new DateTime(2024, 6, 15, 12, 30, 0), (DateTime)rect[1, 0]!);
    }

    // -------------------------------------------------------------------------
    // Embedded nulls — split by direction because the full round-trip is
    // fragile: pandas's float64 dtype turns Arrow nulls into NaN on
    // `to_pandas()`, and `from_pandas()` rendering of NaN can lose the
    // null-bitmap distinction. We test each leg in isolation against a
    // contract that doesn't depend on that conversion.
    // -------------------------------------------------------------------------

    [Fact]
    public void Execute_TableInputWithNulls_PythonSeesAsMissing()
    {
        // C# → kernel direction. Encodes nulls via Arrow null bitmap;
        // pandas in the kernel exposes them as NaN. The script counts
        // NaN/null cells per column via df.isna().sum() so the assertion
        // is type-agnostic: we only require the kernel to see "missing"
        // in the right positions, not a specific Python representation.
        using var fx = new KernelFixture();
        var script = fx.WriteScript("count_nulls.py",
            "def transform(df):\n" +
            "    return df.isna().sum().tolist()\n");

        var result = PyRun.Execute(
            script: script,
            input: new object?[,]
            {
                { 1.0, 10.0 },
                { null, 20.0 },
                { 3.0, null },
            },
            kwargs: null,
            client: fx.Client);

        // Column 0 has 1 null (row 1); column 1 has 1 null (row 2).
        var rect = Assert.IsType<object?[,]>(result);
        Assert.Equal(2, rect.GetLength(0));
        Assert.Equal(1, rect.GetLength(1));
        Assert.Equal(1.0, rect[0, 0]);
        Assert.Equal(1.0, rect[1, 0]);
    }

    [Fact]
    public void Execute_PythonReturnsListWithNone_DecodesAsVectorWithNull()
    {
        // kernel → C# direction. A Python list with None goes through
        // pa.array(...) which uses the Arrow null bitmap (not NaN) for
        // None positions, so the C# decoder sees a clean null.
        using var fx = new KernelFixture();
        var script = fx.WriteScript("list_with_none.py",
            "def transform():\n" +
            "    return [1.0, None, 3.0]\n");

        var result = PyRun.Execute(
            script: script,
            input: null,
            kwargs: null,
            client: fx.Client);

        // Column vector → N×1 rectangle.
        var rect = Assert.IsType<object?[,]>(result);
        Assert.Equal(3, rect.GetLength(0));
        Assert.Equal(1, rect.GetLength(1));
        Assert.Equal(1.0, rect[0, 0]);
        Assert.Null(rect[1, 0]);
        Assert.Equal(3.0, rect[2, 0]);
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
