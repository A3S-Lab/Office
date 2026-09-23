# R2 conditional-format precedence corpus

Permanent Spreadsheet gate for Traditional Office / WPS conditional-format
priority and Stop-if-true behavior.

## Scope

- Higher-priority rules win conflicting text/fill properties.
- Non-conflicting properties from later rules still merge when Stop-if-true is
  off.
- Stop-if-true blocks later rules for matched cells, including different
  format properties.
- Native XLSX export writes ascending `priority` and preserves `stopIfTrue`.

## Run

```bash
bun run test:corpus:r2-cf
```

Large-range worker evaluation, broader CF formula functions, and advanced
visuals remain ordered leftovers beyond this gate.
