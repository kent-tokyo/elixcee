using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using ClosedXML.Excel;

// Persistent JSON-lines worker: JIT and library caches survive batch boundaries.
// All workbook handles and correctness assertions are outside the timed disposal.
if (!OperatingSystem.IsMacOS())
    throw new PlatformNotSupportedException("This worker's durability contract is macOS F_FULLFSYNC.");

if (args.SequenceEqual(new[] { "--self-test" }))
{
    try
    {
        Native.FullSync(-1);
        throw new InvalidOperationException("Invalid descriptor must not sync successfully.");
    }
    catch (Win32Exception)
    {
        Console.WriteLine("sync failure propagation: OK");
    }
    return;
}
if (!args.SequenceEqual(new[] { "--server" }))
    throw new ArgumentException("Use --server or --self-test");

string? line;
while ((line = Console.ReadLine()) != null)
{
    var request = JsonSerializer.Deserialize<Request>(line)
        ?? throw new ArgumentException("Missing request");
    if (request.iterations < 1 || request.iterations > 1000)
        throw new ArgumentOutOfRangeException(nameof(request.iterations));
    using var original = new XLWorkbook(request.fixture);
    if (original.Worksheets.Count != 1)
        throw new ArgumentException("Benchmark requires one sheet");
    var expected = Snapshot.Read(original);
    var sheetName = original.Worksheet(1).Name;
    expected[$"{sheetName}!A1"] = new CellState(null, 123);
    expected[$"{sheetName}!B1"] = new CellState("1+2", default);
    var samples = new List<object>();
    for (var i = 0; i < request.iterations; ++i)
    {
        var watch = Stopwatch.StartNew();
        using var workbook = new XLWorkbook(request.fixture);
        var loaded = watch.Elapsed.TotalMilliseconds;
        var sheet = workbook.Worksheet(1);
        sheet.Cell("A1").Value = 123;
        sheet.Cell("B1").FormulaA1 = "1+2";
        var mutated = watch.Elapsed.TotalMilliseconds;
        var temporary = Path.Combine(Path.GetDirectoryName(request.output)!,
            $".closedxml-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var file = new FileStream(temporary, FileMode.CreateNew,
                       FileAccess.ReadWrite, FileShare.None))
            {
                workbook.SaveAs(file); // Default: no formula recalculation/package validation.
                file.Flush(); // Drain managed buffers, then the exact same OS barrier as Rust.
                Native.FullSync(file.SafeFileHandle.DangerousGetHandle().ToInt32());
                GC.KeepAlive(file);
            }
            Native.Replace(temporary, request.output);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
        var saved = watch.Elapsed.TotalMilliseconds;
        using var reread = new XLWorkbook(request.output);
        var elapsed = watch.Elapsed.TotalMilliseconds;
        var actual = Snapshot.Read(reread);
        if (actual.Count != expected.Count ||
            expected.Any(pair => !actual.TryGetValue(pair.Key, out var value) || value != pair.Value))
            throw new InvalidDataException("Cell value/formula mismatch after round trip");
        samples.Add(new { load_ms = loaded, mutate_ms = mutated - loaded,
            save_ms = saved - mutated, reload_ms = elapsed - saved, total_ms = elapsed });
    }
    Console.WriteLine(JsonSerializer.Serialize(new {
        closedxml = typeof(XLWorkbook).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion,
        runtime = RuntimeInformation.FrameworkDescription,
        architecture = RuntimeInformation.ProcessArchitecture.ToString(),
        file_sync = "F_FULLFSYNC", samples
    }));
}

record Request(string fixture, string output, int iterations);
record CellState(string? Formula, XLCellValue Value);

static class Snapshot
{
    public static Dictionary<string, CellState> Read(XLWorkbook workbook)
    {
        var result = new Dictionary<string, CellState>();
        foreach (var sheet in workbook.Worksheets)
            foreach (var cell in sheet.CellsUsed(XLCellsUsedOptions.Contents))
            {
                if (!cell.HasFormula && cell.CachedValue.IsBlank) continue;
                var key = $"{sheet.Name}!{cell.Address.ToStringRelative()}";
                result.Add(key, cell.HasFormula
                    ? new CellState(cell.FormulaA1.TrimStart('='), default)
                    : new CellState(null, cell.CachedValue));
            }
        return result;
    }
}

static class Native
{
    [DllImport("libSystem.B.dylib", SetLastError = true)]
    private static extern int fcntl(int fd, int command);

    [DllImport("libSystem.B.dylib", SetLastError = true)]
    private static extern int rename(string oldPath, string newPath);

    public static void FullSync(int fd)
    {
        if (fcntl(fd, 51 /* F_FULLFSYNC */) != 0)
            throw new Win32Exception(Marshal.GetLastPInvokeError(), "F_FULLFSYNC failed");
    }

    public static void Replace(string source, string destination)
    {
        if (rename(source, destination) != 0)
            throw new Win32Exception(Marshal.GetLastPInvokeError(), "atomic rename failed");
    }
}
