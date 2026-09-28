# Real-file corpus (Phase 2)

Permanent fidelity gate built from **redistributable** documents authored
outside A3S Office. Synthetic JSZip fixtures stay under `r0-*` / `r2-*`; this
tree proves import → no-op export → reopen on third-party bytes.

## License

| File | Source | License |
| --- | --- | --- |
| `fixtures/SampleDoc.docx` | [apache/poi](https://github.com/apache/poi) `test-data/document/SampleDoc.docx` | Apache-2.0 |
| `fixtures/charts.xlsx` | [apache/poi](https://github.com/apache/poi) `test-data/spreadsheet/123233_charts.xlsx` | Apache-2.0 |
| `fixtures/FormulaEvalTestData.xlsx` | [apache/poi](https://github.com/apache/poi) `test-data/spreadsheet/FormulaEvalTestData_Copy.xlsx` | Apache-2.0 |
| `fixtures/SampleShow.pptx` | [apache/poi](https://github.com/apache/poi) `test-data/slideshow/SampleShow.pptx` | Apache-2.0 |
| `fixtures/PDFBOX-4352-0.pdf` | [Apache PDFBox](https://pdfbox.apache.org/) sample redistributed in [mozilla/pdf.js](https://github.com/mozilla/pdf.js) `test/pdfs/PDFBOX-4352-0.pdf` | Apache-2.0 |

Do not add fixtures without a redistributable license and a row in this table.
Prefer Apache POI, LibreOffice sample collections, or other OSI-licensed sets.

## Contract

For each fixture:

1. `importOfficeFile` succeeds (or fails closed with a listed compatibility
   issue — never a silent empty document).
2. `createArtifactBlob` produces a re-openable package of the same kind.
3. Re-import succeeds; critical package parts that were present on first import
   remain present (part names for OOXML; page count for PDF).
4. Every intentional normalization appears in `compatibility.issues`.

Independent LibreOffice headless reopen (Phase 2.2) is optional locally when
`soffice` is on `PATH`; CI enables it when the runner image provides it.

## Run

```bash
bun run test:corpus:real
# optional, when LibreOffice is installed:
bun run test:corpus:real:libreoffice
```
