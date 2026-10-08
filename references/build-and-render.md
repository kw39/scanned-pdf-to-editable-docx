# Build and render mechanics

Read for a new reconstruction, complex layouts, footer migration, font checking, or renderer setup.

## Source manifest and transcription

A manifest is the comparison reference, not something inferred back from the generated DOCX. Proofread it against the original scan. Use consecutive **physical** page indices; printed page labels can skip numbers. Store the source path, verification status, font evidence, paragraph roles, measured baseline pitches and gaps between blocks, chosen Word spacing settings, images allowed to remain, writing-line counts, footer labels and optional-question rules.

A complete machine-readable example is [examples/manifest.json](../examples/manifest.json). Its values belong to a synthetic test page, not a document preset.

Observed distances describe the scan; chosen settings describe the Word implementation and must be checked in the render. All observed distances are **baseline to baseline** in rendered-page points: block `spacing.observed_baseline_pitch_pt` is the interval between lines within the block; `spacing.observed_baseline_to_next_block_pt` is the distance from the block's last line to the first line of the next listed block on the same page. The latter includes a line of text and is not automatically Word's paragraph After value. Record implementation choices separately as `chosen_word_rule`, `chosen_word_line_pt`, `chosen_before_pt` and `chosen_after_pt` when useful; the checker does not validate those settings.

Page `answer_line_count` and `answer_line_pitch_pt` describe typed writing lines. Pitch requires a count and only one answer-line specification is allowed per physical page. Page-wide pitch applies to every consecutive detected writing line; for distinct answer groups with larger gaps, check regions separately instead. Cover fields and mixed-label lines need visual inspection; the checker detects underscore-only rows of at least ten characters.

A requested next-block distance requires that next block on the same page. List blocks in source reading order and restrict columns with `render_bbox`; this is not a general column-adjacency detector. The checker compares every interval, not just the median, with tolerances `--pitch-tolerance` (default 1 pt) and `--gap-tolerance` (default 3 pt). See [layout-calibration.md](layout-calibration.md#match-line-spacing-and-gaps-to-the-scan) for measurement and verification.

`lines` hold original printed lines within one logical paragraph. `render_bbox` is optional and refers to the **rendered PDF**, in points from its top-left. It is useful for repeated passages, columns and table cells; it is not a raw phone-scan pixel box. Do not silently turn uncertain OCR groups into verified paragraph boundaries. Include every prose block; map other objects and writing spaces separately. A paragraph split over a page boundary needs explicit per-page fragments, with their total line count and split verified.

Inspect whether an existing PDF text layer is trustworthy. Otherwise render at roughly 200–300 dpi, then use available OCR or image transcription. Check English/CJK text, decimals, currencies, units, symbols, equations and marks directly against the image. No OCR service is mandatory. A source with ambiguous content needs clarification; an inferred best-match font must not be labelled metadata recovered from a raster scan.

## Physical paper size versus scan dimensions

Choose a common physical paper size before measuring layout. Record the scan's raster/PDF dimensions separately from the intended Word dimensions; crop boxes and camera perspective are not paper-size instructions. Use one size for pages of the same orientation unless a real size change is clearly evidenced. For a genuine landscape page, normally swap the same paper's width and height. A rotated portrait scan alone is not evidence of a landscape original.

In the manifest, `size_pt` describes the intended Word page; keep differing source dimensions separately as `scan_size_pt`, and explain any genuine page-size exception. For revisions, preserve the user's accepted paper size unless changing it is requested. Normalize source coordinates to the common page and text column before calibrating indents, tables, figures and spacing. Do not stretch fonts or equations to compensate for inconsistent cropping.

Check `w:sectPr/w:pgSz` for every section and verify the exported page dimensions, including pages with unique headers/footers. Creating a section for a page label does not justify changing its paper size.

## Styles, alignment and typography

For a new paper, use named styles for actual roles such as Title, Instructions, Question, Subpart, RomanSubpart, Choice, SourceText, AnswerLine, MarkLine and Footer. Choose roles from this scan, not a preset hierarchy. Put shared typography, paragraph alignment and source-matched line/paragraph spacing in styles; retain local variations that the original demonstrates. For revisions, patch the latest file and preserve its accepted formatting rather than normalising it wholesale.

Read effective formatting from document defaults, the based-on style chain, paragraph/run overrides and relevant language/font slots. Table and numbering styles can contribute formatting too; the audit's alignment resolver does not resolve those conditional styles. Logical start/end alignment remains start/end in the report; its physical edge depends on writing direction. A mixed Normal style can legitimately contain centred titles and numeric cells: its outlier list is a diagnostic, not a command to justify them all.

Choose alignment from source evidence:

- Justified paragraphs: non-final lines share a right edge and word spaces vary. Apply Word Justify (`WD_ALIGN_PARAGRAPH.JUSTIFY`, OOXML `w:jc` value `both`), leaving the last line naturally short.
- Left-aligned paragraphs: ragged right edge and generally even word spaces.
- Centred paragraphs: line midpoints share the intended column centre.
- Right-aligned paragraphs: right edges align while left edges vary.

Use related multi-line paragraphs of the same role when a short or distorted paragraph cannot establish alignment. Do not infer left alignment merely because a final line is short. Verify the effective Word setting as well as the visible edges; equal line counts alone do not establish correct justification. True centring still permits legitimate symmetric column indents; asymmetric hand-tuned indents must not substitute for centre alignment. Resolve an intentional user edit before restoring a source style.

For new OOXML documents targeting modern Word, set the `compatibilityMode` setting to 15 and verify the result. Do not switch an existing user-edited document's mode without rendering all pages again. A mode flag alone neither fixes fonts nor guarantees identical wrapping.

CJK font selection uses `w:rFonts/@w:eastAsia` as well as Latin font slots. Inspect real Chinese glyphs, punctuation compression and line-breaking rules; test mixed Latin/CJK lines. Use ordinary text runs for simple inline mathematics such as `g(4) = 69`, with source-matched italics and native sub/superscripts. Use OMML where layout needs stacked fractions, radicals, matrices or similar mathematical structures. Check token order and sub/superscript relationships in typed maths, and fraction/root/stacked structure in OMML; visible token sequences alone are insufficient. Prefer portable Unicode to legacy Symbol-font mappings where the glyph/font supports it; retain source symbols and verify the render.

## Real footers and section breaks

Use footer/header stories rather than body-anchored boxes. A run such as `P.{PAGE} of 19` can use a PAGE field with the printed total `19` as literal text even if the supplied scan has only 17 physical pages. When the scan skips a printed page, create the necessary section and page-number restart (`w:pgNumType/@w:start`), or use literal per-section labels.

Unlink the applicable default/first/even footer stories when their text differs. Use different-first-page/even-page settings only where the scan requires them. A section break can replace an existing page break; blindly retaining both can create a blank page. Check the trailing paragraph after tables and the footer/body clearance. Migration from floating labels is an intentional structural change: confirm the same printed labels remain on the same physical pages, rather than demanding the old text-box counters remain identical.

## Native diagram placement and other layouts

Native groups, curves, arrows and label text remain editable. Word groups are commonly floating objects; do not assume every Shape group can be converted to an InlineShape. Use an inline drawing container only when supported by the builder. Otherwise use a group/canvas with a single stable anchor near the same question, and suitable paragraph-relative placement; check reflow in Word. Grouping and correct anchoring are distinct checks.

Multi-option diagrams and candidate-number/OMR boxes often fit native borderless grids or small native tables. Check both reading order and border visibility. Two-column pages need genuine columns or deliberate native layout containers and separate column/block comparisons. Landscape pages need appropriate sections. Rotated margin text needs editable rotated labels. Photographs may remain images; formulas, axes and graph labels do not become editable simply because an SVG looks sharp.

If using raw OOXML for drawings/equations, verify Word opens the file normally without repair and that expected objects survive. Inspect modern `mc:Choice` and legacy `mc:Fallback` representations; count one effective object, but do not erase required fallbacks simply to silence an audit warning. Update stale representations consistently.

## Renderer and font evidence

Detect Word availability; do not assume Windows or Word exists. Microsoft Word is preferred for calibrating a Word deliverable. The optional Windows helper is:

```text
powershell -File scripts/render_word.ps1 -Docx result.docx -Pdf qa.pdf -MapJson word_map.json
```

It opens an isolated temporary copy read-only in an empty hidden Word instance, disables macros, exports PDF and a paragraph page map, and never saves the source. Its map uses Word physical start/end page numbers separately from adjusted start/end page numbers. Existing QA outputs are refused rather than overwritten; use a new QA filename for each run. It does not prove that an unreadable source was transcribed correctly.

When Word is unavailable, render with an available engine such as LibreOffice and record its identity. Check fonts in the exported PDF (`pdffonts`, or PDF text/font inspection) against the intended font slots. A metrically compatible substitute can be useful, but it is not the identical typeface and engine differences remain. A substitute or small safety allowance is not proof of how the user's Word will wrap. Report the evidence and actual render scope precisely.

Do not change the user's open Word session or kill all Word processes to resolve a hang. The helper refuses an instance that already contains documents, closes only its own copy, and quits only when the instance is empty. Diagnose a failed render before retrying; external time limits must target only this task's worker. Permission limits do not become authorisation to control an unrelated session.

## Review the verification reports

The audit does not modify its DOCX inputs; optional reports require new filenames. It canonicalises one modern alternate-content branch for counts, retains fallback-only content, warns about stale alternate text, includes header/footer/note stories and mathematical structure, and resolves ordinary paragraph-style alignment. It does not prove glyph fidelity, rendering, section attachment or drawing geometry. Story text comparison retains story types but cannot alone prove the footer-to-page mapping.

Case, punctuation and some Unicode forms are normalized by the PDF line matcher, so compare exact wording separately.

A zero exit code means the reported invariants passed, not that the reconstruction is finished. Preservation differences require reconciliation with authorised changes. Exact underscore sequences can change deliberately during line-length calibration; actual visible writing-line counts must still match. The legacy `underscore_line_count_unchanged` key measures underscore **paragraphs**, not rendered lines.

The audit also flags page-label candidates with explicit page syntax kept in body text boxes (use real footers), paragraphs placed with a negative right indent or a large left indent (use real centre/right alignment or one consistent indent/tab grid), and possible printed-line splits into separate paragraphs. Bare graph labels such as `P1`, `P2` and `0` are not page-label evidence; inspect the reported label snippets. Ambiguous page furniture still needs the manifest/render check. Bare slash labels such as `1/2` or `3/19` are reported separately as `ambiguous_fraction_or_page_labels`, without recommending a footer change. Even a series sharing a denominator can be maths labels; verify its source role and page position. Alignment outliers, text-strip picture heuristics, exact table heights and old compatibility settings need inspection. Native diagrams can contain short images legitimately; Word can keep accurate legacy drawing representations. Avoid treating every warning as an instruction to edit.