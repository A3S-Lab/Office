# R2 formula-compatibility corpus

Permanent Spreadsheet gate for Traditional Office / WPS daily calculation
parity.

## Scope

- Reuses `tests/fixtures/spreadsheet-kernel-parity.json` (kernel scalar
  contracts).
- Adds `fixtures/r2-formula-daily.json` for common daily formulas (`SUM`,
  `AVERAGE`, `IF`, `COUNT`, `AND`/`OR`, `ROUND`/`POWER`/`SQRT`/`ABS`,
  `CONCAT`/`IFERROR`/`MOD`, `SUBTOTAL`/`COUNTA`, `LEFT`/`RIGHT`/`LEN`/`MID`,
  `DATE`/`YEAR`/`MONTH`/`DAY`, `SUMIF`/`COUNTIF`/`VLOOKUP`,
  `SUMIFS`/`COUNTIFS`, `AVERAGEIF`/`AVERAGEIFS`, `INDEX`/`MATCH`/`HLOOKUP`,
  approximate ascending `VLOOKUP`/`MATCH`/`HLOOKUP`, `SUMPRODUCT`/`PRODUCT`/
  `VALUE`, bounded `NUMBERVALUE` with explicit separators, locale-guessing
  `VALUE`, bounded `SEQUENCE` / `TRANSPOSE` / `UNIQUE` spill) plus fail-closed
  unsupported and volatile functions (`TODAY`/`NOW`/`RAND`).

## Run

```bash
bun run test:corpus:r2
```

Pivot/slicer depth, remaining spill functions (`FILTER`/`SORT`),
comma literals, and Fortune-owned redraw replacement remain ordered R2 leftovers
beyond this bootstrap.
