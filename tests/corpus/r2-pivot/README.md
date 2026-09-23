# R2 pivot-depth corpus

Permanent Spreadsheet gate for Traditional Office / WPS pivot authoring depth.

## Scope

- Built-in pivot-style catalog (`WORK_SPREADSHEET_PIVOT_STYLE_OPTIONS`).
- Fail-closed diagnostics for calculated cache fields
  (`xlsx.pivots.calculated-fields`) vs grouping (`xlsx.pivots.grouping`).
- Fail-closed slicer/timeline package detection (`xlsx.pivots.slicers`).
- Editable slicer UI, timelines, and pivot charts remain leftovers beyond this
  gate.

## Run

```bash
bun run test:corpus:r2-pivot
```
