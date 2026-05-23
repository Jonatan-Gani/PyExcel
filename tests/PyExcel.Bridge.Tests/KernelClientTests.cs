using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using PyExcel.Bridge;
using PyExcel.Kernel.Client;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// End-to-end tests for <see cref="KernelClient"/> against a real Python
/// kernel subprocess. Runs on Linux and Windows — the Python transport
/// handles both backends.
///
/// <para>The tests deliberately avoid encoding Arrow payloads on the C#
/// side: <c>tests/kernel/test_arrow_io.py</c> already exercises the
/// shape-preserving Arrow roundtrip in full. What these tests verify is
/// the C# side of the wire — request-meta construction, frame dispatch,
/// reply-meta parsing, exception mapping. Scripts therefore take no
/// positional args (kwargs go through canonical JSON, not Arrow) and the
/// response payloads are checked for presence/absence rather than decoded
/// contents.</para>
/// </summary>
public class KernelClientTests
{
    // -------------------------------------------------------------------------
    // Argument validation (no kernel needed)
    // -------------------------------------------------------------------------

    [Fact]
    public void Constructor_Null_Supervisor_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => new KernelClient(null!));
    }

    [Fact]
    public void Run_Null_Request_Throws()
    {
        using var fixture = new KernelFixture();
        var client = new KernelClient(fixture.Supervisor);
        Assert.Throws<ArgumentNullException>(() => client.Run(null!));
    }

    [Fact]
    public void Run_Empty_Script_Throws()
    {
        using var fixture = new KernelFixture();
        var client = new KernelClient(fixture.Supervisor);
        Assert.Throws<ArgumentException>(() =>
            client.Run(new RunRequest { Script = "" }));
    }

    [Fact]
    public void Cancel_Rejects_Empty_RunId()
    {
        using var fixture = new KernelFixture();
        var client = new KernelClient(fixture.Supervisor);
        Assert.Throws<ArgumentException>(() => client.Cancel(""));
        Assert.Throws<ArgumentException>(() => client.Cancel(null!));
    }

    // -------------------------------------------------------------------------
    // Happy path: success returns RunResult with the right shape
    // -------------------------------------------------------------------------

    [Fact]
    public void Run_Returns_Payload_When_Function_Returns_Value()
    {
        using var fixture = new KernelFixture();
        var script = fixture.WriteScript("const.py",
            "def transform():\n    return 42\n");

        var client = new KernelClient(fixture.Supervisor);
        var result = client.Run(new RunRequest { Script = script });

        Assert.False(result.IsEmpty);
        Assert.Single(result.Payloads);
        Assert.True(result.Payload.Length > 0, "Arrow payload should be non-empty");
        Assert.True(result.DurationMs >= 0);
    }

    [Fact]
    public void Run_Returns_Empty_Payloads_When_Function_Returns_None()
    {
        using var fixture = new KernelFixture();
        var script = fixture.WriteScript("none.py",
            "def transform():\n    return None\n");

        var client = new KernelClient(fixture.Supervisor);
        var result = client.Run(new RunRequest { Script = script });

        Assert.True(result.IsEmpty);
        Assert.Empty(result.Payloads);
        Assert.Throws<InvalidOperationException>(() => _ = result.Payload);
    }

    [Fact]
    public void Run_Passes_Kwargs_Through_Canonical_Json()
    {
        using var fixture = new KernelFixture();
        // Use kwargs to encode a value the script can spit back out via a
        // side-channel file. Avoids needing Arrow encode on the C# side.
        var sentinel = Path.Combine(fixture.ScratchDir, "kwargs_seen.txt");
        var script = fixture.WriteScript("kw.py",
            "def transform(*, factor, label, sentinel):\n" +
            "    with open(sentinel, 'w') as f:\n" +
            "        f.write(f'{label}:{factor}')\n" +
            "    return None\n");

        var client = new KernelClient(fixture.Supervisor);
        client.Run(new RunRequest
        {
            Script = script,
            Kwargs = new Dictionary<string, object?>
            {
                ["factor"] = 5L,
                ["label"] = "doubled",
                ["sentinel"] = sentinel,
            },
        });

        Assert.True(File.Exists(sentinel), "kwargs were never received");
        Assert.Equal("doubled:5", File.ReadAllText(sentinel));
    }

    [Fact]
    public void Run_Echoes_Custom_RunId()
    {
        using var fixture = new KernelFixture();
        var script = fixture.WriteScript("id.py",
            "def transform():\n    return 1\n");

        var client = new KernelClient(fixture.Supervisor);
        var result = client.Run(new RunRequest
        {
            Script = script,
            RunId = "my-custom-run-id",
        });
        Assert.Equal("my-custom-run-id", result.RunId);
    }

    [Fact]
    public void Run_Auto_Generates_RunId_When_Omitted()
    {
        using var fixture = new KernelFixture();
        var script = fixture.WriteScript("auto.py",
            "def transform():\n    return 1\n");

        var client = new KernelClient(fixture.Supervisor);
        var result = client.Run(new RunRequest { Script = script });
        Assert.False(string.IsNullOrEmpty(result.RunId));
        Assert.True(result.RunId.Length >= 16, "auto-generated run_id looks too short");
    }

    [Fact]
    public void Run_Honours_Custom_Function_Name()
    {
        using var fixture = new KernelFixture();
        var sentinel = Path.Combine(fixture.ScratchDir, "which_function.txt");
        var script = fixture.WriteScript("custom.py",
            "def transform():\n" +
            $"    open(r'{sentinel}', 'w').write('default'); return 0\n" +
            "def my_func():\n" +
            $"    open(r'{sentinel}', 'w').write('custom'); return 0\n");

        var client = new KernelClient(fixture.Supervisor);
        client.Run(new RunRequest { Script = script, Function = "my_func" });

        Assert.Equal("custom", File.ReadAllText(sentinel));
    }

    [Fact]
    public void Run_Sequential_Calls_Reuse_Kernel()
    {
        using var fixture = new KernelFixture();
        var counter = Path.Combine(fixture.ScratchDir, "count.txt");
        var script = fixture.WriteScript("counter.py",
            "_calls = [0]\n" +
            "def transform():\n" +
            "    _calls[0] += 1\n" +
            $"    open(r'{counter}', 'w').write(str(_calls[0]))\n" +
            "    return _calls[0]\n");

        var client = new KernelClient(fixture.Supervisor);
        for (var i = 1; i <= 3; i++)
        {
            client.Run(new RunRequest { Script = script });
            Assert.Equal(i.ToString(), File.ReadAllText(counter));
        }
        // The module-level _calls survives across runs → same kernel + cached
        // module, exactly the persistent-supervisor promise.
    }

    // -------------------------------------------------------------------------
    // Error paths
    // -------------------------------------------------------------------------

    [Fact]
    public void Run_Propagates_User_Exception_As_KernelException()
    {
        using var fixture = new KernelFixture();
        var script = fixture.WriteScript("boom.py",
            "def transform():\n    raise ValueError('no good')\n");

        var client = new KernelClient(fixture.Supervisor);
        var ex = Assert.Throws<KernelException>(() =>
            client.Run(new RunRequest { Script = script }));

        Assert.Equal("Exception", ex.Code);
        Assert.Equal("ValueError", ex.PythonType);
        Assert.Contains("no good", ex.Message);
        Assert.Contains("boom.py", ex.PythonTraceback);
        Assert.True(ex.DurationMs >= 0);
    }

    [Fact]
    public void Run_Missing_Script_Yields_ModuleNotFound()
    {
        using var fixture = new KernelFixture();
        var client = new KernelClient(fixture.Supervisor);

        var ex = Assert.Throws<KernelException>(() =>
            client.Run(new RunRequest { Script = "/definitely/not/a/real/file.py" }));

        Assert.Equal("ModuleNotFound", ex.Code);
    }

    [Fact]
    public void Run_Function_Not_Found_Yields_FunctionNotFound()
    {
        using var fixture = new KernelFixture();
        var script = fixture.WriteScript("noxform.py",
            "def something_else():\n    return 1\n");

        var client = new KernelClient(fixture.Supervisor);
        var ex = Assert.Throws<KernelException>(() =>
            client.Run(new RunRequest { Script = script }));

        Assert.Equal("FunctionNotFound", ex.Code);
    }

    // -------------------------------------------------------------------------
    // Async wrapper
    // -------------------------------------------------------------------------

    [Fact]
    public async Task RunAsync_Happy_Path_Returns_Result()
    {
        using var fixture = new KernelFixture();
        var script = fixture.WriteScript("async.py",
            "def transform():\n    return 'ok'\n");

        var client = new KernelClient(fixture.Supervisor);
        var result = await client.RunAsync(new RunRequest { Script = script });

        Assert.Single(result.Payloads);
    }

    [Fact]
    public async Task RunAsync_Surfaces_KernelException()
    {
        using var fixture = new KernelFixture();
        var script = fixture.WriteScript("async_boom.py",
            "def transform():\n    raise RuntimeError('async-failure')\n");

        var client = new KernelClient(fixture.Supervisor);
        var ex = await Assert.ThrowsAsync<KernelException>(() =>
            client.RunAsync(new RunRequest { Script = script }));

        Assert.Equal("Exception", ex.Code);
        Assert.Equal("RuntimeError", ex.PythonType);
        Assert.Contains("async-failure", ex.Message);
    }

    // -------------------------------------------------------------------------
    // Fixture: spawns one kernel subprocess + a scratch dir, cleans both up.
    // -------------------------------------------------------------------------

    private sealed class KernelFixture : IDisposable
    {
        public KernelSupervisor Supervisor { get; }
        public string ScratchDir { get; }

        public KernelFixture()
        {
            var python = DiscoverPython()
                ?? throw new InvalidOperationException("no python on PATH");
            var pythonPath = DiscoverEmbeddedPath()
                ?? throw new InvalidOperationException("embedded/ not found near test binary");

            ScratchDir = Path.Combine(
                Path.GetTempPath(),
                "pyexcel-clienttest-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(ScratchDir);

            Supervisor = KernelSupervisor.StartPython(python, pythonPath);
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

    // -------------------------------------------------------------------------
    // Discovery helpers — duplicated from KernelSupervisorTests.
    // -------------------------------------------------------------------------

    private static string? DiscoverPython()
    {
        var candidates = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? new[] { "python.exe", "python3.exe" }
            : new[] { "python3", "python" };
        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (var dir in pathEnv.Split(Path.PathSeparator))
        {
            if (string.IsNullOrWhiteSpace(dir)) continue;
            foreach (var name in candidates)
            {
                var full = Path.Combine(dir, name);
                if (File.Exists(full)) return full;
            }
        }
        return null;
    }

    private static string? DiscoverEmbeddedPath()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        for (var i = 0; i < 8 && dir != null; i++)
        {
            var candidate = Path.Combine(dir.FullName, "embedded", "pyexcel", "kernel", "__main__.py");
            if (File.Exists(candidate))
                return Path.Combine(dir.FullName, "embedded");
            dir = dir.Parent;
        }
        return null;
    }
}
