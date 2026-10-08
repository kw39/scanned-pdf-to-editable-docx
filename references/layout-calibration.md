# Layout calibration and paragraph fit

Read when Word fits fewer words per line than the scan, alignment looks wrong, or reflow shifts questions and answer spaces.

## Layout priorities

1. Match the source font and preferred line-ending words.
2. If line endings differ but the paragraph keeps the same number of lines in the same font, accept that wrap. Do not distort an acceptable paragraph just to chase an ending.
3. If the paragraph gains a line, recheck font/size, column width, indents and mark space, then slightly narrow that paragraph if necessary. If it loses a line, investigate the block's calibration instead of adding padding or arbitrary breaks.

The strict requirement is the printed line count of each paragraph plus its page allocation and source-matched vertical spacing, not equal total line counts over a whole page. A paragraph gaining a line and its neighbour losing one is not a pass.

## Calibrate globally, then check paragraphs

Use source lines from several pages: question stems, ordinary and nested subparts, source passages, table cells and final lines with marks. Measure word/glyph widths relative to their actual column; phone-scan scale and perspective vary. Compare glyph shapes to candidate fonts and measure size per block. A font substitution, wrong font slot, or incorrect table padding must be corrected before compensating with character condensation.

Start from the chosen common physical paper size, natural font metrics, usable width and role styles. Normalize cropped scans to this geometry; do not vary Word page sizes to chase scan boundaries or line endings. Match the source's paragraph alignment before tuning fit. Justification adjusts word spaces; it does not inherently fit more words, and a correct line count does not excuse left alignment when the source is justified. If needed, adjust character scale/spacing in small measured steps (for example 1% scale or 0.05 pt spacing). For ordinary prose, usually stay within about 95–100% scale and 0 to −0.2 pt spacing, with visual checks; these are diagnostic guardrails, not universal font presets. Do not assume 90% scale or −0.5 pt spacing is invisible. If more is needed, revisit the font/width estimate and whether the scan used a condensed face. Do not substitute another face merely to force a fit; record measured adjustments before proceeding.

Keep equations, CJK font slots and answer-line metrics separate from Latin prose calibration. Never split flowing prose into a paragraph per printed line. Do not force every line ending with soft breaks; that conceals wrong metrics and can stretch justified lines.

## Match line spacing and gaps to the scan

Source fidelity takes priority over visual polish. Check internal line spacing and gaps between blocks separately: prose can have the correct baseline pitch while the page still looks too dense because the space before a subpart or around a table was removed. Compare at the same page/column scale; screenshots at different zoom levels do not establish point values.

Record representative vertical landmarks on each source page (question/subpart starts, table top/bottom, first answer line and last body line) and compare their rendered positions. This catches cumulative drift that isolated line-count checks miss. Use paragraph Before/After and native table spacing/placement to reproduce source gaps, rather than filling spare space evenly or making every gap uniform.

Measure baseline-to-baseline distance across several ordinary lines of each role, rather than the white gap between glyph bounding boxes. Convert source measurements to page points using the page/column dimensions; correct scan skew and local photo perspective before comparing distances. Record the observed baseline pitch and gaps between paragraphs, subparts, choices and questions. Include source passages, table cells and answer-line pitch; preserve deliberate differences rather than imposing one spacing on the whole paper.

Choose Word line spacing and paragraph Before/After settings that reproduce those measurements. A raster scan does not reveal whether the original used Single, Multiple, At least or Exactly, or how a block gap was implemented: record the measured target separately from the chosen Word settings. Use shared role styles for a new paper, and patch the latest accepted user formatting when revising. Do not insert repeated blank paragraphs or reduce line spacing to compensate for wrong wrapping or page allocation. Do not treat 1.15 or 1.5 as a default for reconstruction.

Exactly spacing can clip tall equations, accents or superscripts. For paragraphs containing stacked fractions, radicals, equation arrays/matrices or tall scripts, use At least with the measured prose pitch as the minimum, or natural Single/Multiple spacing. This allows the equation-bearing line to grow; do not apply an undersized fixed line box merely to match the ordinary prose pitch. Resolve inherited spacing as well as direct formatting. Preserve the existing equation size and structure when repairing clearance; do not shrink or rasterize maths to hide the defect. If the extra height shifts later content, recalibrate source-supported paragraph gaps and verify question starts, writing lines, footer clearance and page count. After rendering, compare successive baselines, first/last text-line positions and gaps between blocks, as well as writing-line pitch and footer clearance. Matching paragraph line counts alone does not establish matching line spacing. Measure the scan with `scripts/measure_scan_vertical.py`. Give the scale as the rendered width of one full justified line (`--column-pt`) and the scan width of the same line in pixels (`--column-px`), and a `--box` around one column of printed text that excludes figures. It reports steps between detected ink-band centres, which approximate baseline pitch only for comparable glyph bands; these are measurements to confirm on the source image, not recovered typographic baselines. The rendered checker likewise uses full-height glyph bottoms as baseline proxies, so mixed font sizes or equations require visual/coordinate verification. A text-line step near the line pitch is a line inside a paragraph; a larger step includes a block gap. It also reports ruled lines (answer lines and table borders alike). Cross-check the scale against the page height: phone scans are not always exactly proportioned, and a difference of a few percent is uncertainty to record. Write the confirmed values into the manifest `spacing` fields and `answer_line_count`/`answer_line_pitch_pt`; `compare_line_endings.py` then compares them with the render, checking every declared pitch interval so a matching median cannot hide an uneven line. Keep intentional paragraph gaps outside an intra-paragraph pitch check. Page-wide answer pitch is suitable for one uniformly spaced run; check separate runs and mixed-label answer fields individually. Handwriting on answer lines adds noise, so confirm every value on the page image. If a distorted scan prevents a reliable estimate, record that uncertainty rather than claiming an exact measurement.

## Compare paragraph references, not guessed blocks

Prefer the proofread manifest:

```text
python scripts/compare_line_endings.py render.pdf --manifest scan_manifest.json --json fit.json
```

Alternatively write a UTF-8 transcription with explicit paragraph boundaries:

```text
#page 1
#block q2b
(b) The original first printed line
and its continuation line.

#block q3stem
3. The next original paragraph.
```

Then use `--scan-lines transcription.txt --verified-reference`. `--scan-tsv page1.tsv page2.tsv` uses OCR paragraph groups; add `--verified-reference` only after the agent checks every relevant transcription and paragraph boundary against the images. File order defines physical page order. Without verified boundaries, the output is provisional and cannot pass delivery.

The checker keeps Unicode/CJK tokens, requires a full unambiguous match on the intended physical page, and compares line counts separately for each explicit paragraph ID. Repeated text or complex columns need `render_bbox` in the manifest. Missing/ambiguous text, missing pages, empty references and unverified inputs cannot give a success verdict. A matched paragraph with the same count but different endings is acceptable.

The tool checks only listed blocks. A contiguous match does not prove that no extra wording surrounds it; compare the complete Word paragraphs/transcription separately. Confirm complete source coverage, printed page labels, writing lines, alignment and diagrams separately. Fractions, stacked mathematics, rotated text and complex table geometry may require direct render/structure checks; PDF text-line extraction does not by itself establish their visual line layout.

Rendered lines are grouped by vertical overlap, so real subscripts and superscripts (S₁, x², 10³) count as part of their printed line. An unresolved block reports `differences_first_5` (reference tokens against the closest rendered tokens) and, when its text sits on another rendered page, `found_on_other_render_pages`. Use them to correct the transcription, the DOCX wording or the page allocation; neither field turns the block into a pass. Similarity is a bounded 0–1 heuristic, not confidence in the transcription; check the scan before changing either text.

Do not narrow text in response to an OCR false flag. Compare the source image, correct the transcription/reference, and rerun first. Do not infer paragraph boundaries from shared line breaks: many are ordinary internal lines, and equal differences can cancel across unrelated paragraphs.

## Words hyphenated across printed lines

Preserve the printed `consump-` / `tion` lines in the source manifest and record the proofread full word `consumption` separately. Distinguish a discretionary line-end hyphen from a real lexical hyphen (such as `cost-effective`); do not remove every hyphen or concatenate arbitrary words. In the editable paragraph use the complete word and, when appropriate for the scan, a native optional/soft hyphen rather than a forced return.

The line checker matches printed hyphens strictly by default. A lower-case continuation is not proof that the hyphen is discretionary: `re-` / `sign` may mean `re-sign`, not `resign`. After the agent has proofread the complete word against the source, annotate only the confirmed breaks in that block's JSON manifest:

```json
{"id": "q4", "page": 1, "lines": ["Explain the consump-", "tion of goods."],
 "discretionary_hyphen_breaks": [1]}
```

Indices are 1-based, referring to the preceding nonempty printed line. This agent annotation does not require user approval. Unannotated lexical or uncertain breaks retain their hyphens. Text/TSV inputs can match as printed; use a proofread manifest to authorise optional-hyphen variants.

For confirmed breaks the checker tries the printed form, omitted hyphen, then joined fragments. It reports the matching `hyphenation_variant` and confirmed break indices. A joined word has different endings and passes only if the paragraph's line count matches. Other hyphens in the paragraph remain strict. A block that still does not match stays incomplete; verify wording, visible line count, page allocation and hyphenation against the scan and Word map/render, and record the check. The agent completes routine checks rather than handing them back to the user.

## Marks, answer spaces and vertical flow

Reserve room for a right-tabbed mark. If the source places it separately, use a right-aligned mark paragraph kept with its question, or a tested soft-break treatment. For an intentional soft break, Word supports `<w:doNotExpandShiftReturn w:val="1"/>` under `w:compat` in `word/settings.xml`. It suppresses expansion of the preceding justified line: verified with compatibility mode 15 in Word 16.0. Test the intended renderer and existing settings rather than applying it blindly. Separate mark paragraphs avoid depending on that behaviour.

Native underscore lines contribute to paragraph flow. Give them stable font metrics, enough width to avoid wrapping, and deliberate line spacing. Count visible writing lines in the render, not just underscore paragraphs. Changing prose line counts shifts tables, the first/last writing line and footer clearance; verify those positions after every layout-sensitive correction.

Use explicit page/section breaks where the source starts a new page, avoiding chains of unnecessary blank paragraphs or Keep-with-next settings. End-of-section labels and page totals belong in the area the original uses. Missing printed pages must not generate invented blank pages.

## Common fidelity failures

- Lost gaps before subparts or around tables make a page too dense even when its text fits. Measure and restore the source gaps.
- Sentence pictures, scanned boxes and image-based marks/answer lines need native replacements.
- Assuming left alignment or applying justification indiscriminately both lose source fidelity.
- Diagram options require native axes, curves, arrowheads and labels with handwriting excluded.
- Blank paragraphs and keep-with-next chains can move question starts; body-anchored page labels can duplicate footers.
- Incorrect prose pitch, block gaps and answer-line count can partly cancel so a page still fits. Check each measurement and vertical landmark separately.
- A one-line source paragraph reconstructed as two lines fails even if neighbouring paragraphs compensate. A paragraph retaining its source line count can accept different endings.

## Equation clearance regression check

A Word reconstruction used Exactly 21 pt for a paragraph containing an inline fraction. The denominator existed in OMML and PDF text extraction, while part of it was visibly clipped in Word. Whole-page review and passing prose line counts missed the defect. The repair changed only equation-bearing paragraph line rules to At least, keeping their existing minimum heights. The value 21 pt is evidence of this failure, not a safe/unsafe threshold for other equations.

Run `audit_docx.py` and review `layout_lints.equation_clipping_risks`: it resolves direct, paragraph-style/basedOn and document-default spacing and flags tall OMML under Exactly. It does not measure actual glyph visibility, resolve conditional table/numbering styles or certify At least as sufficient; fixed table rows, text-box extents, overlapping objects and other constraints still need checking.

After any equation or line-spacing edit, render in the intended Word renderer and inspect **every tall equation at close zoom**, including complete numerator and denominator strokes/descenders, fraction bars, radical tips, stacked rows, delimiters and script edges. Inspect equation lines separately from prose line-count checks. Inspect the affected full pages for overlap and displaced answer spaces, then verify page allocation across the document. Record the equation checks explicitly before delivery; native text/equation counts or a successful PDF text extraction cannot substitute for visible glyph clearance. When the symptom occurs only in Word's editing view, reproduce and check that view if available, and report any verification limitation.
