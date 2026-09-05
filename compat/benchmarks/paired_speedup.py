"""Interleaved before/after comparison with exact decompressed ZIP equality."""
import argparse
import json
import platform
import statistics
import subprocess
import tempfile
import time
import zipfile
from pathlib import Path

import openpyxl
from equal_workbook_speed import check, digest, logical
from compare_closedxml import summarize


def zip_parts(path):
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        assert len(names) == len(set(names)), 'duplicate ZIP member'
        return {name: archive.read(name) for name in names}


def compare_parts(before, after):
    left, right = zip_parts(before), zip_parts(after)
    assert left.keys() == right.keys(), 'ZIP member list changed'
    for name in left:
        assert left[name] == right[name], f'ZIP member bytes changed: {name}'


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--before', type=Path, required=True)
    parser.add_argument('--after', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--pairs', type=int, default=30)
    parser.add_argument('--mixed', action='store_true', help='also check a mixed numeric/Unicode text fixture')
    args = parser.parse_args()
    if args.pairs < 10 or args.pairs % 2:
        parser.error('pairs must be even and at least 10')
    repo = Path(__file__).resolve().parents[2]
    binaries = {k:v.resolve(strict=True) for k,v in [('before',args.before), ('after',args.after)]}
    metadata = dict(date=time.strftime('%Y-%m-%d'), platform=platform.platform(),
        python=platform.python_version(), openpyxl=openpyxl.__version__,
        rustc=subprocess.check_output(['rustc','--version'],text=True).strip(),
        binary_sha256={k:digest(v) for k,v in binaries.items()},
        after_source_sha256={p:digest(repo/p) for p in ['src/lib.rs', 'src/reader.rs',
            'src/vm/mod.rs', 'Cargo.lock', 'examples/bench_workbook.rs',
            'compat/benchmarks/paired_speedup.py']},
        pairs=args.pairs, warmups_per_engine_per_fixture=3, target_speedup=1.2,
        timing='load/edit/standard durable save/rename/reload; worker startup/assertions excluded',
        validation='all decompressed ZIP members byte-identical each pair; openpyxl all-cell checks each ten pairs',
        fixtures=[])
    with tempfile.TemporaryDirectory(prefix='elixcee-speedup-pairs-') as directory:
        directory = Path(directory)
        fixtures = [('small', repo/'tests/fixtures/e2e/source.xlsx')]
        for rows in (1000, 10000):
            path = directory/f'numeric-{rows}x10.xlsx'
            wb = openpyxl.Workbook()
            wb.active.title = 'Sheet1'
            for row in range(1, rows+1):
                wb.active.append([row*10+col for col in range(1, 11)])
            wb.save(path)
            wb.close()
            fixtures.append((f'numeric-{rows}x10', path))
        if args.mixed:
            path = directory/'mixed-1000x10.xlsx'
            wb = openpyxl.Workbook()
            wb.active.title = 'Sheet1'
            for row in range(1, 1001):
                wb.active.append([row*10+col if col % 3 == 1 else
                    f' 共通 {row % 17} &<> sheetViews ' if col % 3 == 2 else
                    f'固有 {row}:{col} 日本語' for col in range(1, 11)])
            wb.save(path)
            wb.close()
            fixtures.append(('mixed-1000x10', path))
        for name, fixture in fixtures:
            outputs = {k:directory/f'{k}.xlsx' for k in binaries}
            expected = logical(fixture)
            assert len(expected) == 1
            next(iter(expected.values())).update(A1=123, B1='=1+2')
            samples = {k:[] for k in binaries}

            def run(key):
                result = subprocess.run([str(binaries[key]), str(fixture), str(outputs[key]), '1'],
                    capture_output=True, text=True, check=True, timeout=120)
                payload = json.loads(result.stdout)
                assert len(payload['samples']) == 1
                return payload['samples'][0]

            for _ in range(3):
                for key in binaries:
                    run(key)
            orders = []
            for pair in range(args.pairs):
                order = ['before', 'after'] if pair % 2 == 0 else ['after', 'before']
                orders.append(order)
                for key in order:
                    samples[key].append(dict(pair=pair+1, **run(key)))
                compare_parts(outputs['before'], outputs['after'])
                if (pair+1) % 10 == 0 or pair+1 == args.pairs:
                    for output in outputs.values():
                        check(output, expected)
            summary = {k:summarize(v) for k,v in samples.items()}
            ratio = summary['before']['total_ms']['p50']/summary['after']['total_ms']['p50']
            paired_ratios = [b['total_ms']/a['total_ms'] for b,a in zip(samples['before'],samples['after'])]
            result = dict(fixture=name, fixture_sha256=digest(fixture),
                summaries=summary, p50_speedup=ratio, median_paired_speedup=statistics.median(paired_ratios),
                target_met=ratio >= 1.2, zip_parts_verified=True,
                output_bytes={k:v.stat().st_size for k,v in outputs.items()}, orders=orders, samples=samples)
            metadata['fixtures'].append(result)
            print(name, json.dumps({k:v for k,v in result.items() if k not in ('samples','orders')}), flush=True)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(metadata, indent=2)+'\n')


if __name__ == '__main__':
    main()
