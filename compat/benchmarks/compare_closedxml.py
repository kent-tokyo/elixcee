"""Three-library, equal-durability macOS benchmark; run after release builds."""
import argparse
import json
import math
import os
import platform
import select
import statistics
import subprocess
import sys
import tempfile
import tomllib
from pathlib import Path

import openpyxl
from equal_workbook_speed import check, digest, logical, py_batch


class ClosedXmlWorker:
    def __init__(self, dotnet, dll):
        env = dict(os.environ, DOTNET_CLI_TELEMETRY_OPTOUT='1',
                   DOTNET_ROOT=str(dotnet.parent), DOTNET_TieredCompilation='0')
        self.errors = tempfile.TemporaryFile(mode='w+t')
        self.process = subprocess.Popen([str(dotnet), str(dll), '--server'],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=self.errors,
            text=True, env=env)
        self.metadata = None

    def run(self, fixture, output, count):
        request = dict(fixture=str(fixture), output=str(output), iterations=count)
        self.process.stdin.write(json.dumps(request)+'\n')
        self.process.stdin.flush()
        ready, _, _ = select.select([self.process.stdout], [], [], 180)
        if not ready:
            raise TimeoutError('ClosedXML batch exceeded 180 seconds')
        line = self.process.stdout.readline()
        if not line:
            self.errors.seek(0)
            raise RuntimeError('ClosedXML failed: '+self.errors.read())
        result = json.loads(line)
        assert result['file_sync'] == 'F_FULLFSYNC'
        self.metadata = {k:v for k,v in result.items() if k != 'samples'}
        return result['samples']

    def close(self):
        self.process.stdin.close()
        try:
            self.process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            self.process.kill()
            self.process.wait()
        self.process.stdout.close()
        self.errors.close()


def summarize(samples):
    result = {}
    for metric in ('load_ms', 'mutate_ms', 'save_ms', 'reload_ms', 'total_ms'):
        values = sorted(s[metric] for s in samples)
        assert values and all(math.isfinite(v) and v >= 0 for v in values)
        result[metric] = dict(p50=statistics.median(values),
                             p95=values[math.ceil(.95*len(values))-1])
    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--elixcee', type=Path, required=True)
    parser.add_argument('--dotnet', type=Path, required=True)
    parser.add_argument('--closedxml', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--rounds', type=int, default=3)
    parser.add_argument('--iterations', type=int, default=10)
    args = parser.parse_args()
    if sys.platform != 'darwin':
        parser.error('F_FULLFSYNC worker currently supports macOS only')
    if args.rounds < 3 or args.rounds % 3 or not 5 <= args.iterations <= 1000:
        parser.error('rounds must be a positive multiple of 3; iterations must be 5..1000')
    repo = Path(__file__).resolve().parents[2]
    binary, dotnet, dll = (p.resolve(strict=True) for p in
                           (args.elixcee, args.dotnet, args.closedxml))
    project = repo/'compat/benchmarks/closedxml'
    manifest = tomllib.loads((repo/'Cargo.toml').read_text())
    worker = ClosedXmlWorker(dotnet, dll)
    results = []
    try:
        with tempfile.TemporaryDirectory(prefix='elixcee-closedxml-bench-') as root:
            root = Path(root)
            fixtures = [('small', repo/'tests/fixtures/e2e/source.xlsx')]
            for rows in (1000, 10000):
                fixture = root/f'numeric-{rows}x10.xlsx'
                wb = openpyxl.Workbook()
                wb.active.title = 'Sheet1'
                for row in range(1, rows+1):
                    wb.active.append([row*10+col for col in range(1, 11)])
                wb.save(fixture)
                wb.close()
                fixtures.append((f'numeric-{rows}x10', fixture))
            for name, fixture in fixtures:
                expected = logical(fixture)
                assert len(expected) == 1
                next(iter(expected.values())).update(A1=123, B1='=1+2')
                output = root/'result.xlsx'
                samples = {k:[] for k in ('elixcee', 'closedxml', 'openpyxl')}
                sizes = {}

                def run(engine, count):
                    if engine == 'elixcee':
                        completed = subprocess.run([str(binary), str(fixture), str(output), str(count)],
                            capture_output=True, text=True, check=True, timeout=180)
                        payload = json.loads(completed.stdout)
                        assert payload['version'] == manifest['package']['version']
                        observations = payload['samples']
                    elif engine == 'closedxml':
                        observations = worker.run(fixture, output, count)
                    else:
                        observations = py_batch(fixture, output, count, expected)
                    assert len(observations) == count
                    check(output, expected)  # Independent, whole-output read after every batch.
                    assert not list(root.glob('.closedxml-*.xlsx')), 'unpublished temp file leaked'
                    sizes[engine] = output.stat().st_size
                    return observations

                for engine in samples:
                    run(engine, 5)
                orders = []
                for round_no in range(args.rounds):
                    order = list(samples)
                    offset = round_no % len(order)
                    order = order[offset:] + order[:offset]
                    orders.append(order)
                    for engine in order:
                        observations = run(engine, args.iterations)
                        samples[engine].extend(dict(round=round_no+1, **s) for s in observations)
                summaries = {k:summarize(v) for k,v in samples.items()}
                results.append(dict(fixture=name, input_sha256=digest(fixture),
                    input_bytes=fixture.stat().st_size, output_bytes=sizes, verified=True,
                    orders=orders, summaries=summaries, samples=samples))
                print(name, json.dumps(summaries), flush=True)
    finally:
        worker.close()

    source_files = ['src/lib.rs', 'src/reader.rs', 'src/vm/mod.rs', 'Cargo.lock',
        'examples/bench_workbook.rs', 'compat/benchmarks/equal_workbook_speed.py',
        'compat/benchmarks/compare_closedxml.py', 'compat/benchmarks/closedxml/Program.cs',
        'compat/benchmarks/closedxml/ClosedXmlBench.csproj',
        'compat/benchmarks/closedxml/global.json', 'compat/benchmarks/closedxml/packages.lock.json']
    payload = dict(schema_version=1, platform=platform.platform(),
        python=platform.python_version(), openpyxl=openpyxl.__version__, openpyxl_lxml=openpyxl.LXML,
        elixcee=manifest['package']['version']+'-unreleased-working-tree',
        closedxml=worker.metadata, dotnet_tiered_compilation=False,
        dotnet_sdk=subprocess.check_output([str(dotnet), '--version'],cwd=project,text=True).strip(),
        rustc=subprocess.check_output(['rustc','--version'],text=True).strip(),
        head=subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo,text=True).strip(),
        source_sha256={p:digest(repo/p) for p in source_files},
        elixcee_binary_sha256=digest(binary),
        closedxml_artifact_sha256={p.name:digest(p) for p in sorted(dll.parent.iterdir())
                                  if p.suffix in ('.dll', '.json')},
        rounds=args.rounds, iterations_per_round=args.iterations, warmup_per_engine_per_fixture=5,
        file_sync='F_FULLFSYNC', atomic_rename=True, directory_sync=False,
        timing='load + A1/B1 edit + save + file barrier + rename + reload; startup/assertions/disposal excluded',
        validation='all cell values/formula strings; no full OOXML or formula-result equivalence claim',
        fixtures=results)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(payload, indent=2)+'\n')


if __name__ == '__main__':
    main()
