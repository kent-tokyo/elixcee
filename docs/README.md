# Documentation map

Documentation version: **1.0.7**. API/coverage pages describe this version;
dated benchmark records retain their original baseline and source hashes.
The JavaScript package remains private at 0.0.0-development.
The current branch may contain unreleased changes; the published 1.0.7
contract ends at the `[1.0.7]` section of the changelog.

elixcee is a headless Excel workbook automation runtime. Its native contract
covers workbook editing, supported formula recalculation, and data-processing
VBA execution/diagnosis; it is not a VBA-only runner or a complete Excel desktop
replacement.

## Find the right document

| Need | Maintained reference |
|---|---|
| Install and first example | [English](../README.md), [日本語](../README_ja.md), [中文](../README_zh.md) |
| Browser-first beginner workflow | [Playground](../playground/README.md) with an in-page English/Japanese/Simplified Chinese switch |
| Current priorities and remaining gates | [Roadmap](../ROADMAP.md) |
| Roadmap dependency classification | [Dependency audit](roadmap-dependency-audit-2026-09-10.md) |
| VBA and formula coverage | [FUNCTIONS](../FUNCTIONS.md) |
| Formula argument/mode contracts | [formula-contracts.json](../compat/formula-contracts.json); validate with `python3 -B scripts/check-formula-dispatch.py --check-docs --check-contracts` |
| Python signatures | [elixcee.pyi](../elixcee.pyi) |
| CLI commands and JSON schema | [Agent contract](agent-contract.md) |
| Support guarantees and non-goals | [v1 support contract](v1-support-contract.md) |
| Migration examples and boundaries | [Migration guide](migration.md) |
| Resource ceilings and sizing evidence | [Limits](limits.md) |
| Input/output threat model | [Security](xlsx-security-model.md) |
| Reader/VM/writer and JS boundaries | [Architecture](xlsx-architecture.md) |
| OOXML feature status matrix | [Feature matrix](../compat/ooxml-feature-matrix.json) |
| JS runtime / types / known differences | [Package](../packages/xlsx/README.md), [types](typescript-compatibility.md), [differences](compatibility-known-defects.md) |
| JS compatibility target | [Compatibility goal](xlsx-compatibility-goal.md) |
| Dependency licensing | [Licensing](licensing.md), [notices](../THIRD_PARTY_NOTICES.md) |
| Changes by release | [1.x / Unreleased](../CHANGELOG.md), [0.x archive](history/CHANGELOG-0.x.md) |

## Evidence, not current API specifications

- [Workbook benchmarks](benchmarks/README.md): fixed-library comparisons and
  incremental speedups, including withdrawn comparisons and unmet targets.
- [Reader/VM measurements](measurements/README.md): dated calibration, raw samples,
  platform limits, and unfinished validation.
- [Compatibility harnesses](../compat/README.md): oracle, differential, and fixture tests.

Results apply only to their recorded source/binary, input, host, and timing policy.
Do not replace a historical result with a current-sounding claim or multiply
speedup ratios across different baselines.

The current Chart evidence is deliberately split: bounded creation and editing
are BUILD-stage capabilities, while the macOS Excel reopen gate has only passed
for a line-chart-only smoke case; the multi-chart/bar case still repairs.

## Maintenance rules

Keep commands/schema in the agent contract, coverage in FUNCTIONS, limits in limits,
and open work in the roadmap. Link to those pages instead of copying their full text.
Date measurements and distinguish implemented, locally tested, measured, and published.

Versioned designs and exact-name compatibility pointers are retained when source,
tests, or evidence still reference them. Local ignored internal_docs notes are not
required to understand the public contract and may be absent from a fresh clone.
