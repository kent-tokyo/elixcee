# Equal-workbook benchmark: elixcee, ClosedXML, and openpyxl

Date: 2026-09-11 (Asia/Tokyo)  
Host: macOS arm64  
elixcee: 1.0.10 current source release binary  
ClosedXML: 0.105.1 (commit `b4ebe47bd3ecebf8480dea7422188d16b35b9d9a`)  
.NET: 10.0.400 / runtime 10.0.11  
openpyxl: 3.1.2

Each library loaded the same workbook, changed A1 and B1, saved with
`F_FULLFSYNC`, atomically renamed the result, and reloaded it. Five warmups and
three rounds of ten measured iterations were used. The output was independently
checked for cell values and formula strings after every batch.

| Fixture | elixcee p50 | ClosedXML p50 | openpyxl p50 | elixcee vs ClosedXML | elixcee vs openpyxl |
|---|---:|---:|---:|---:|---:|
| small | 5.033 ms | 10.265 ms | 12.578 ms | 2.04× faster | 2.50× faster |
| 1,000×10 | 15.775 ms | 131.007 ms | 142.630 ms | 8.30× faster | 9.04× faster |
| 10,000×10 | 112.434 ms | 1,052.784 ms | 1,370.149 ms | 9.36× faster | 12.19× faster |

These are same-host workflow timings, not universal library rankings. The
benchmark excludes startup and disposal, does not measure peak RSS, and does
not claim formula-result or full OOXML equivalence. The complete raw report is
[workbook-closedxml-2026-09-11.json](workbook-closedxml-2026-09-11.json).

Reproduction:

```sh
/Users/k_tanabe/.dotnet/dotnet build --configuration Release --no-restore \
  compat/benchmarks/closedxml/ClosedXmlBench.csproj
cargo build --release --example bench_workbook --offline
python3 compat/benchmarks/compare_closedxml.py \
  --elixcee target/release/examples/bench_workbook \
  --dotnet /Users/k_tanabe/.dotnet/dotnet \
  --closedxml compat/benchmarks/closedxml/bin/Release/net10.0/ClosedXmlBench.dll \
  --rounds 3 --iterations 10 \
  --output /tmp/elixcee-closedxml-comparison.json
```
