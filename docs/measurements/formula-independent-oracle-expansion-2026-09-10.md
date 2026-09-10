# Formula oracle expansion

Date: 2026-09-10 (Asia/Tokyo)

The formula-only oracle runner now probes additional statistical and
distribution functions, including `CORREL`, `SLOPE`, `INTERCEPT`, `RSQ`,
`FISHER`, `FISHERINV`, `GAMMA`, `GAMMALN`, `NORM.S.*`, `COVARIANCE.P`,
`CHISQ.DIST.RT`, `F.DIST`, `T.DIST.2T`, and scalar projections of multi-column
`LINEST`／`LOGEST`／`TREND`／`GROWTH` results.

## Local result

LibreOffice 26.2.5 on this host evaluated 118 of 155 probes, with 118/118
matches. Thirty-seven probes were excluded from the comparable count because this
LibreOffice build returned an unsupported result in the formula-only path.

The current-source `elixcee` 1.0.6 wheel, built locally for CPython 3.13 on
arm64 macOS and installed in an isolated environment, evaluated the same 112
comparable probes with 112/112 matches. This is a source-build verification,
not a published-package or cross-platform claim.

The six additional comparable probes cover two multi-column `LINEST`
coefficients, its R-squared row, one multi-column `LOGEST` factor, and one
prediction each from multi-column `TREND` and `GROWTH`; all six matched
LibreOffice's result. The paired installed-wheel rerun for this
expanded fixture is still pending because the current shell does not have the
`elixcee` module installed.

The seven newly added distribution probes were retained as visible probes;
LibreOffice returned `#NAME?` for them here. The newly added `CORREL`,
`SLOPE`, `INTERCEPT`, `RSQ`, `FISHER`, `FISHERINV`, and `GAMMALN` probes were
comparable and matched their expected values.

The additional boundary probes for `NORM.*`, `GAMMA.*`, `BETA.*`,
`WEIBULL.DIST`, `EXPON.DIST`, `LOGNORM.*`, `F.DIST.2T`, and `T.DIST.*` were
also retained as visible probes; this LibreOffice build returned `#NAME?` for
these names in the formula-only path.

The conversion/integer probes `GCD`, `LCM`, `QUOTIENT`, `ROMAN`,
`NETWORKDAYS.INTL`, `EVEN`, and `ODD` were comparable and matched. This build
returned `#NAME?` for the retained `ARABIC`, `BASE`, and `DECIMAL` probes.

`SUMSQ`, `DEVSQ`, `TRIMMEAN`, `CONVERT`, and `DOLLARDE` also matched. The
`MODE.SNGL` name is not evaluated by this LibreOffice path. Its `DOLLARFR`
notation differs from the Excel-compatible contract used by elixcee, so it is
kept outside the cross-engine count rather than being reported as a match.

With the runner's finite-number tolerance, `GEOMEAN`, `HARMEAN`, `AVEDEV`,
`SKEW`, and a four-value `KURT` sample also matched. Text and error results
remain exact comparisons.

The direct-coercion probes `SUM("2",TRUE)` and `AVERAGE("2",TRUE)` are kept
visible but excluded because this LibreOffice build returns `#VALUE!` for both
forms. The implementation follows Excel's documented distinction: `SUM`
coerces direct text/logical arguments, while `AVERAGE` ignores them and returns
`#DIV/0!` when no numeric argument remains.

This is independent LibreOffice evidence plus a paired local source-build
check. It is not Microsoft Excel compatibility evidence, and it does not claim
that the same result holds for every published wheel or platform.

## Reproduction

```bash
python3 compat/oracle/run-libreoffice-formulas.py \
  --soffice /opt/homebrew/bin/soffice \
  --output /tmp/elixcee-formula-probe-20260910.json
```

After installing a locally built wheel into an isolated environment, add
`--with-elixcee` to run the paired 118-case check.

The runner exits non-zero if any comparable LibreOffice probe mismatches.
