# Scanned PDF to Editable Word

An agent skill for reconstructing scanned exams, worksheets and printed forms as editable Word documents while matching the original layout. Supports English, Chinese, bilingual papers and mathematical content, plus revisions that preserve edits in the latest user-supplied DOCX.

This is an agent workflow with verification helpers. It does not include a one-command OCR-to-DOCX converter, an OCR model, or a universal document builder. The agent chooses available transcription, authoring and rendering tools.

## Install and use

Copy this repository into a folder named `scanned-pdf-to-editable-docx` in your agent's skill directory. Keep `SKILL.md`, `agents/`, `references/`, `scripts/`, `examples/`, `README.md`, `LICENSE` and `requirements.txt` together. Developer tests and GitHub configuration are optional at runtime.

Example request:

> Use $scanned-pdf-to-editable-docx to reconstruct this scanned paper as an editable DOCX, matching its printed layout and omitting handwritten answers.

For revisions, provide both the scan and the latest edited DOCX and describe the changes you want preserved. Explicit user instructions override the skill's defaults.

The default destination is an `output/` directory beside the source PDF. Existing results are preserved with numbered filenames. Native prose, tables, equations and simple drawing objects remain editable; warranted photographs and logos may remain images. Handwriting is omitted by default. Paragraph layout, page labels, answer spaces and equation clearance are checked against the scan.

## Requirements

The Python helpers target **Python 3.10 or newer**. Choose an available local or host-managed runtime. For the optional layout helpers:

```sh
python -m pip install -r requirements.txt
```

- `audit_docx.py`: Python standard library only.
- `compare_line_endings.py`: pdfplumber; requires a rendered PDF with a text layer.
- `measure_scan_vertical.py`: Pillow and numpy; works on a rendered scan image.
- `render_word.ps1`: Windows PowerShell 5.1 and installed desktop Microsoft Word. Word automation is optional; another available renderer can be used with its limitations recorded.

The agent also needs tools for PDF/image inspection, DOCX authoring and rendering. OCR is optional and must be proofread. No paid service or particular OCR provider is required. Dependencies are installed separately; fonts and renderers are not bundled.

## Verification helpers

Run commands from the repository root. Replace example input paths with your own files; create the QA directory first. JSON report paths must be new files, so use a new filename for each run.

```sh
python scripts/audit_docx.py result.docx --report qa/audit.json
python scripts/audit_docx.py result.docx --baseline latest-edited.docx --report qa/preservation.json
python scripts/compare_line_endings.py qa/render.pdf --manifest scan_manifest.json --json qa/fit.json
python scripts/measure_scan_vertical.py qa/page.png --column-pt 475 --column-px 1583 --box 100 200 1683 1000
```

On Windows, optional Word rendering:

```powershell
powershell -NoProfile -File scripts/render_word.ps1 -Docx result.docx -Pdf qa/render.pdf -MapJson qa/word-map.json
```

The Word helper uses a temporary read-only copy, disables macros, and refuses existing outputs or an instance containing open documents. It does not save the input or terminate unrelated Word processes. COM automation can still fail or hang; diagnose it and limit cleanup to the task's own worker.

| Helper | Exit 0 | Exit 1 | Exit 2 | Exit 3 |
| --- | --- | --- | --- | --- |
| DOCX audit | Reported invariants passed; warnings may remain | Input/output or parsing error | Baseline preservation difference | Warnings with `--strict` |
| PDF line comparison | Listed checks passed | Input/output or parsing error | Line count, page count or spacing mismatch | Incomplete/unverified reference or measurement |
| Scan measurement | Measurements produced | Input or measurement error | Argument usage error | — |

Argument parsing may also return 2 for invalid command usage. Deliberate preservation differences must be reconciled with the requested changes. Audit warnings are review candidates. Scan measurements are estimates, not a pass/fail fidelity check.

See [SKILL.md](SKILL.md), [build and render guidance](references/build-and-render.md), and [layout calibration](references/layout-calibration.md). [examples/manifest.json](examples/manifest.json) is a complete synthetic reference for the generated PDF used by the tests; it is not a measurement preset for other papers.

## Limits of the checks

A successful script does not certify a faithful conversion. The line matcher normalizes case, punctuation and some Unicode forms; it checks only listed blocks and declared measurements, and cannot establish exact wording, source coverage, editability of every object, glyph visibility or diagram geometry. The audit inspects structure and selected preservation invariants, not all formatting or footer attachment. Compare every rendered page to its scan and inspect tall equations at close zoom before delivery. Word and other renderers can wrap text differently; record the engine and fonts actually checked.

## Development

```sh
python -m pip install -r requirements-dev.txt
python -m unittest discover -s tests -v
```

Tests generate synthetic DOCX packages, PDFs and images in temporary directories. Private papers, handwritten material and source scans are not included. GitHub Actions runs these checks on Linux and Windows across Python 3.10 and 3.12, with an additional Windows PowerShell syntax check. To run the optional desktop Word integration test on Windows with Word installed:

```powershell
$env:SKILL_TEST_WORD = '1'
python -m unittest discover -s tests -p test_word_render.py -v
```

It creates only a synthetic temporary document and checks export, page mapping, source preservation, cleanup status and refusal to overwrite outputs. The test limits its worker to 60 seconds; if automation times out, inspect and clean up only that worker's temporary document/instance before retrying.

The workflow follows [GitHub's Python testing guidance](https://docs.github.com/en/actions/tutorials/build-and-test-code/python). CI does not run desktop Word or certify visual fidelity.

## License

[MIT](LICENSE). The license covers this skill's code and documentation; it does not grant rights to third-party papers, fonts or images supplied for conversion.
