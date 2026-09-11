# ClosedXML benchmark worker

Local measurement helper only; not an elixcee runtime dependency or a shipped
Cargo artifact. Requires macOS ARM64 for the measured setup, .NET SDK 10.0.400,
and the pinned ClosedXML 0.105.1 NuGet dependency. Transitive versions/content
hashes are captured in `packages.lock.json`.

From this directory (so `global.json` selects the SDK):

```sh
dotnet restore --locked-mode
dotnet build --configuration Release --no-restore
dotnet bin/Release/net10.0/ClosedXmlBench.dll --self-test
```

Put the pinned SDK's `dotnet` executable on PATH, or replace `dotnet` in these
commands with its absolute path.
From the repository root:

```sh
cargo build --release --example bench_workbook --offline
python3 compat/benchmarks/compare_closedxml.py \
  --elixcee target/release/examples/bench_workbook \
  --dotnet dotnet \
  --closedxml compat/benchmarks/closedxml/bin/Release/net10.0/ClosedXmlBench.dll \
  --rounds 3 --iterations 10 \
  --output /tmp/elixcee-closedxml-comparison.json
python3 -m unittest discover -s compat/benchmarks -p 'test_*.py'
```

Python 3.11+ and openpyxl 3.1.2 are needed for the driver. The worker is a
persistent JSON-lines server: each request gives `fixture`, `output`, and
`iterations`; each response gives version/runtime metadata and timing samples.
The driver disables .NET tiered compilation (`DOTNET_TieredCompilation=0`),
uses five excluded warmups per fixture/engine, and leaves normal GC enabled.
Startup/JIT compilation warmed by these operations is not a measured stage;
compilation or GC that still occurs inside a timed call remains included.

All libraries edit A1 and B1 and use file `F_FULLFSYNC`, then same-directory
atomic rename, then reload. The C# worker deliberately refuses other operating
systems rather than pretending a different durability primitive is equivalent.
Neither the file's containing directory nor formula calculated values are part
of the durability/semantic comparison contract.

ClosedXML uses its normal `SaveAs(Stream)` path. No formula recalculation is
requested, and snapshot verification uses `CachedValue` plus `FormulaA1` so it
does not accidentally trigger recalculation. Original/reloaded workbook data
and validation allocations are retained outside the timer; this is **not** a
peak-memory or end-to-end process-lifetime benchmark. An independent openpyxl
read after every batch rejects value/formula loss.

The driver never publishes results externally. See the dated public report in
[ClosedXML report](../../../docs/benchmarks/workbook-closedxml-2026-09-06.md) for the measurements and limits.
