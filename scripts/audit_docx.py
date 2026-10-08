#!/usr/bin/env python3
"""Read-only structural audit for scanned-PDF DOCX reconstructions."""
import argparse
from collections import Counter
import difflib
import hashlib
import json
from pathlib import Path
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
V = "{urn:schemas-microsoft-com:vml}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
# (6 marks) (1 mark) (6marks) [6 marks] (Total: 10 marks) (2 + 3 marks) （6分）
MARK = re.compile(r"[(\[]\s*(?:total\s*:?\s*)?\d+(?:\s*\+\s*\d+)*\s*marks?\s*[)\]]"
                  r"|[（(]\s*\d+\s*分\s*[）)]", re.I)
SECTION_HEADING = re.compile(r"\s*section\b", re.I)
STORY_PART = re.compile(r"word/(header|footer)\d*\.xml$|word/(footnotes|endnotes)\.xml$")
GRAPHIC_KINDS = {
    "http://schemas.openxmlformats.org/drawingml/2006/picture": "pictures",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape": "shapes",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup": "groups",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas": "canvases",
    "http://schemas.openxmlformats.org/drawingml/2006/chart": "charts",
}
EMU_PER_CM = 360000
# Page-furniture candidates need explicit page syntax. Bare P1/P2 are price/curve labels,
# not evidence of a footer, just as bare 0/10 may be axis values.
PAGE_LABEL = re.compile(r"\b(?:p\.\s*|page\s*)\d+\b|\bp\s*\d+\s*(?:of|/)\s*\d+\b"
                        r"|\b\d+\s*of\s*\d+\s*$|^\s*[-\u2013\u2014]\s*\d{1,3}\s*[-\u2013\u2014]\s*$"
                        r"|\u7b2c\s*\d+\s*\u9801", re.I)
BARE_FRACTION = re.compile(r"^\s*\d+\s*/\s*\d+\s*$")
TERMINAL = tuple(".?!:;)]\u3002\uff1f\uff01\uff09\u300d")


def twips(ppr, attribute):
    ind = ppr.find(W + "ind") if ppr is not None else None
    value = ind.get(W + attribute) if ind is not None else None
    if value is None and ind is not None:
        value = ind.get(W + {"left": "start", "right": "end"}.get(attribute, attribute))
    try:
        return int(value) if value is not None else 0
    except ValueError:
        return 0


def compact(text):
    return re.sub(r"\s+", "", text)


def normalise_mark(label):
    return " ".join(unicodedata.normalize("NFKC", label).split())


def visible_text(root):
    result = []
    for item in root.iter():
        if item.tag in (W + "t", M + "t"):
            result.append(item.text or "")
        elif item.tag == W + "sym":
            result.append("<sym %s %s>" % (item.get(W + "font"), item.get(W + "char")))
    return "".join(result)


def strip_fallbacks(root):
    """Count one modern branch; retain a fallback when no Choice exists.

    This is an audit view for modern Word, not a rewrite of the source package.
    Unsupported readers and stale fallback text still require review.
    """
    stale = 0
    for parent in list(root.iter()):
        for child in list(parent):
            if child.tag != MC + "AlternateContent":
                continue
            choices = child.findall(MC + "Choice")
            fallbacks = child.findall(MC + "Fallback")
            selected = choices[0] if choices else (fallbacks[0] if fallbacks else None)
            if choices:
                for fallback in fallbacks:
                    stale += compact(visible_text(choices[0])) != compact(visible_text(fallback))
            for branch in list(child):
                if branch is not selected:
                    child.remove(branch)
    return stale


def equation_signature(node):
    """Ignore run styling while retaining mathematical structure and mathematical properties."""
    if node.tag in (W + "rPr", M + "rPr"):
        return None
    attrs = tuple(sorted((k, v) for k, v in node.attrib.items()
                         if k.startswith(M) or node.tag == W + "sym"))
    return (node.tag, attrs, node.text if node.tag == M + "t" else None,
            tuple(x for child in node if (x := equation_signature(child)) is not None))


def parse_story(xml_bytes):
    root = ET.fromstring(xml_bytes)
    stale = strip_fallbacks(root)
    parents = {child: parent for parent in root.iter() for child in parent}

    def inside(node, tag):
        while node in parents:
            node = parents[node]
            if node.tag == tag:
                return True
        return False

    def paragraph_text(node):
        result = []
        for item in node.iter():
            if inside(item, W + "txbxContent"):
                continue
            if item.tag in (W + "t", M + "t"):
                result.append(item.text or "")
            elif item.tag == W + "tab":
                result.append("\t")
            elif item.tag in (W + "br", W + "cr"):
                result.append("\n")
            elif item.tag == W + "noBreakHyphen":
                result.append("-")
            elif item.tag == W + "sym":
                result.append("<sym %s %s>" % (item.get(W + "font"), item.get(W + "char")))
        return "".join(result)

    elements = [p for p in root.iter(W + "p") if not inside(p, W + "txbxContent")]
    paragraphs = [paragraph_text(p) for p in elements]
    floating = [visible_text(box) for box in root.iter(W + "txbxContent")]
    return root, paragraphs, floating, stale, inside, elements


JC_NAMES = {"both": "justify"}  # Keep logical start/end; physical edges depend on text direction.


def load_paragraph_styles(package, names):
    """styleId -> (basedOn, jc, display name); plus the default paragraph style and docDefaults jc."""
    styles, default_style, default_jc = {}, None, None
    if "word/styles.xml" not in names:
        return styles, default_style, default_jc
    root = ET.fromstring(package.read("word/styles.xml"))
    jc = root.find(f"{W}docDefaults/{W}pPrDefault/{W}pPr/{W}jc")
    default_jc = jc.get(W + "val") if jc is not None else None
    for style in root.findall(W + "style"):
        if style.get(W + "type") != "paragraph":
            continue
        sid = style.get(W + "styleId")
        based, jc, name = style.find(W + "basedOn"), style.find(f"{W}pPr/{W}jc"), style.find(W + "name")
        styles[sid] = (based.get(W + "val") if based is not None else None,
                       jc.get(W + "val") if jc is not None else None,
                       name.get(W + "val") if name is not None else sid)
        if style.get(W + "default") in ("1", "true", "on"):
            default_style = sid
    return styles, default_style, default_jc


def resolve_alignment(p, styles, default_style, default_jc):
    """Direct jc, else the style chain, else docDefaults, else left. Table-style jc is not resolved."""
    ppr = p.find(W + "pPr")
    style_ref = ppr.find(W + "pStyle") if ppr is not None else None
    sid = style_ref.get(W + "val") if style_ref is not None else default_style
    name = styles.get(sid, (None, None, sid or "(none)"))[2]
    direct = ppr.find(W + "jc") if ppr is not None else None
    value = direct.get(W + "val") if direct is not None else None
    seen = set()
    while value is None and sid in styles and sid not in seen:
        seen.add(sid)
        based, value, _ = styles[sid]
        sid = based
    value = value or default_jc or "left"
    return JC_NAMES.get(value, value), name


def drawing_report(root):
    counts = Counter()
    strips, page_like = 0, 0
    for drawing in root.iter(W + "drawing"):
        data = drawing.find(".//" + A + "graphicData")
        kind = GRAPHIC_KINDS.get(data.get("uri") if data is not None else None, "other_graphics")
        counts[kind] += 1
        extent = drawing.find(".//" + WP + "extent")
        if kind == "pictures" and extent is not None:
            cx, cy = int(extent.get("cx", 0)), int(extent.get("cy", 0))
            if cy and cy < 0.9 * EMU_PER_CM and cx / cy > 6:
                strips += 1  # one text line tall and much wider: likely a sentence picture
            if cx > 14 * EMU_PER_CM and cy > 19 * EMU_PER_CM:
                page_like += 1
    return counts, strips, page_like


def equation_clearance_risks(root, style_root):
    """Find tall native maths constrained by effective Exactly paragraph spacing.

    This is a review warning, not a glyph visibility measurement. Resolve each
    spacing attribute through direct formatting, basedOn styles and docDefaults.
    Conditional table/numbering styles are not resolved; inspect those visually.
    """
    styles, default_style = {}, None
    defaults = None
    if style_root is not None:
        defaults = style_root.find(f"{W}docDefaults/{W}pPrDefault/{W}pPr/{W}spacing")
        for style in style_root.findall(W + "style"):
            if style.get(W + "type") == "paragraph":
                styles[style.get(W + "styleId")] = style
                if style.get(W + "default") in ("1", "true", "on"):
                    default_style = style.get(W + "styleId")
    tall = {M + tag for tag in ("f", "rad", "eqArr", "m", "nary", "limLow", "limUpp",
                                "sSub", "sSup", "sSubSup", "sPre", "groupChr")}
    risks = []
    for index, paragraph in enumerate(root.iter(W + "p"), 1):
        if not any(node.tag in tall for math in paragraph.iter(M + "oMath") for node in math.iter()):
            continue
        ppr = paragraph.find(W + "pPr")
        ref = ppr.find(W + "pStyle") if ppr is not None else None
        sid = ref.get(W + "val") if ref is not None else default_style
        chain = [(ppr.find(W + "spacing") if ppr is not None else None, "direct")]
        seen = set()
        while sid in styles and sid not in seen:
            seen.add(sid)
            style = styles[sid]
            chain.append((style.find(f"{W}pPr/{W}spacing"), "style:" + sid))
            based = style.find(W + "basedOn")
            sid = based.get(W + "val") if based is not None else None
        chain.append((defaults, "docDefaults"))
        def effective(attribute, fallback):
            return next(((spacing.get(W + attribute), source) for spacing, source in chain
                         if spacing is not None and spacing.get(W + attribute) is not None),
                        (fallback, "Word default"))
        rule, source = effective("lineRule", "auto")
        if rule == "exact":
            line, _ = effective("line", None)
            risks.append({"paragraph_index": index, "text": " ".join(visible_text(paragraph).split())[:100],
                          "line_rule": rule, "rule_source": source,
                          "line_twips": line, "reason": "Tall equation in a fixed line box may be clipped."})
    return risks


def read_docx(path):
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        root, paragraphs, floating, stale, inside, elements = parse_story(package.read("word/document.xml"))
        styles = load_paragraph_styles(package, names)
        style_root = ET.fromstring(package.read("word/styles.xml")) if "word/styles.xml" in names else None
        media = sorted(n for n in names if n.startswith("word/media/") and not n.endswith("/"))
        other_stories = Counter()
        for name in sorted(n for n in names if STORY_PART.match(n)):
            _, story_paras, story_floating, story_stale, _, _ = parse_story(package.read(name))
            stale += story_stale
            text = compact("".join(story_paras + story_floating))
            if text:
                other_stories[(STORY_PART.match(name).group(1) or STORY_PART.match(name).group(2), text)] += 1
        settings = ET.fromstring(package.read("word/settings.xml")) if "word/settings.xml" in names else None
    mode = next((n.get(W + "val") for n in settings.iter(W + "compatSetting")
                 if n.get(W + "name") == "compatibilityMode"), None) if settings is not None else None
    compat_mode = int(mode) if mode and mode.isdigit() else None

    def is_line(p):
        return re.fullmatch(r"_+", p.strip()) is not None

    answer_lines = [p.strip() for p in paragraphs if is_line(p)]
    prose = [p for p in paragraphs if p.strip() and not is_line(p)]
    marks = [normalise_mark(m) for p in prose if not SECTION_HEADING.match(p)
             for m in MARK.findall(p)]
    text = compact("".join(prose))
    counts, strips, page_like = drawing_report(root)

    # Alignment: resolved per paragraph so the agent can compare each paragraph role with the scan.
    # Group by style AND context: table cells (numbers, headers) legitimately differ from body text.
    aligned = []
    for t, el in zip(paragraphs, elements):
        alignment, style = resolve_alignment(el, *styles)
        aligned.append((t, alignment, style + (" [table cell]" if inside(el, W + "tc") else "")))
    prose_aligned = [(t, a, s) for t, a, s in aligned if t.strip() and not is_line(t)]
    by_style = {}
    for _, a, s in prose_aligned:
        by_style.setdefault(s, Counter())[a] += 1
    outliers = [{"style": s, "alignment": a, "style_majority": by_style[s].most_common(1)[0][0],
                 "text": " ".join(t.split())[:60]}
                for t, a, s in prose_aligned if a != by_style[s].most_common(1)[0][0]]
    justified_lines = sum(1 for t, a, _ in aligned if is_line(t) and a in ("justify", "distribute"))

    equation_risks = equation_clearance_risks(root, style_root)
    warnings = []
    if equation_risks:
        warnings.append(f"{len(equation_risks)} tall-equation paragraphs use effective Exactly line spacing; "
                        "use At least or natural spacing and inspect all numerator/denominator and stack edges at close zoom.")
    if outliers:
        warnings.append(f"{len(outliers)} prose paragraphs differ from the usual alignment of their style; "
                        "check each against the scan.")
    if justified_lines:
        warnings.append(f"{justified_lines} underscore answer-line paragraphs are justified; "
                        "keep answer lines out of prose justification.")
    if media and sum(len(p) for p in prose) < 200:
        warnings.append("Very little native body text with embedded images; inspect for page or sentence pictures.")
    if strips >= 3:
        warnings.append(f"{strips} pictures are one text line tall and much wider; inspect for sentence pictures.")
    if page_like:
        warnings.append(f"{page_like} pictures are roughly page-sized; inspect for page pictures.")
    if stale:
        warnings.append(f"{stale} mc:Fallback copies differ from their mc:Choice text; inspect and update alternate representations consistently.")
    if compat_mode is not None and compat_mode < 15:
        warnings.append(f"compatibilityMode={compat_mode}: Word opens this in Compatibility Mode with older layout rules.")
    revisions = sum(1 for n in root.iter() if n.tag in
                    (W + "ins", W + "del", W + "moveFrom", W + "moveTo", W + "rPrChange", W + "pPrChange", W + "sectPrChange", W + "tblPrChange"))
    if not prose:
        warnings.append("No native body prose found; an empty audit cannot certify a reconstructed paper.")
    exact_rows = sum(1 for n in root.iter(W + "trHeight") if n.get(W + "hRule") == "exact")
    if exact_rows:
        warnings.append(f"{exact_rows} table rows have exact heights; inspect for clipped text.")
    if revisions:
        warnings.append("Tracked revisions are present; the visible wording needs review.")
    pictures = sum(1 for _ in root.iter(A + "blip")) + sum(1 for _ in root.iter(V + "imagedata"))
    equations = [equation_signature(n) for n in root.iter(M + "oMath")]
    full_text = compact("".join(paragraphs))
    lints = {"tall_equations_with_exact_spacing": len(equation_risks),
             "equation_clipping_risks": equation_risks,
             "exact_height_table_rows": exact_rows,
             "manual_break_heavy_paragraphs": sum(1 for p in elements
                if sum(1 for n in p.iter(W + "br") if n.get(W + "type") in (None, "textWrapping")) >= 3),
             "space_padding_paragraphs": sum(1 for t in paragraphs if re.search(r" {3,}", t))}
    # Defects seen in a real accepted result: page labels in body text boxes, paragraphs placed by
    # hand-tuned indents (negative right indent, large left indent imitating centring), and one
    # paragraph per printed line. All are review candidates, not automatic defects.
    body_page_labels = [visible_text(box).strip()[:80] for box in root.iter(W + "txbxContent")
                        if PAGE_LABEL.search(visible_text(box).strip())]
    ambiguous_ratios = [visible_text(box).strip()[:80] for box in root.iter(W + "txbxContent")
                        if BARE_FRACTION.fullmatch(visible_text(box).strip())]
    hand_positioned, split_lines, previous = [], [], None
    for t, el, (_, alignment, style) in zip(paragraphs, elements, aligned):
        lint_text = " ".join(t.split())
        in_body = "[table cell]" not in style
        ppr = el.find(W + "pPr")
        left, right = twips(ppr, "left"), twips(ppr, "right")
        hanging = twips(ppr, "hanging") or twips(ppr, "firstLine")
        if (in_body and lint_text and not is_line(t) and alignment not in ("center", "right", "end") and "\t" not in t
                and not hanging and (right < 0 or (left >= 1440 and len(lint_text) <= 80))):
            hand_positioned.append(lint_text[:60])
        if (in_body and previous and lint_text and not is_line(t) and lint_text[0].islower()
                and not previous.endswith(TERMINAL) and not MARK.search(previous[-14:])):
            split_lines.append(lint_text[:60])
        previous = lint_text if in_body and lint_text and not is_line(t) else (previous if not lint_text else None)
    if body_page_labels:
        warnings.append(f"{len(body_page_labels)} body text boxes look like page labels; use real Word footer/header stories.")
    if hand_positioned:
        warnings.append(f"{len(hand_positioned)} paragraphs are placed with a negative right indent or a large left indent; "
                        "centred text needs real centre alignment and indented text one consistent indent/tab grid. "
                        "Check them against the scan.")
    if split_lines:
        warnings.append(f"{len(split_lines)} paragraphs start in lower case after an unfinished paragraph; "
                        "a printed line may have been split into its own paragraph.")
    lints["body_text_box_page_labels"] = len(body_page_labels)
    lints["body_text_box_page_labels_first_10"] = body_page_labels[:10]
    lints["ambiguous_fraction_or_page_labels"] = len(ambiguous_ratios)
    lints["ambiguous_fraction_or_page_labels_first_10"] = ambiguous_ratios[:10]
    lints["hand_positioned_paragraphs"] = len(hand_positioned)
    lints["hand_positioned_paragraphs_first_10"] = hand_positioned[:10]
    lints["possible_line_split_paragraphs"] = len(split_lines)
    lints["possible_line_split_paragraphs_first_10"] = split_lines[:10]
    floating_anchors = [{"anchor": " ".join(t.split())[:80], "label": visible_text(box)[:100]}
                        for t, p in zip(paragraphs, elements) for box in p.iter(W + "txbxContent")]
    summary = {
        "schema_version": 2,
        "file": str(Path(path).resolve()),
        "body_and_table_paragraphs": len(paragraphs),
        "native_prose_characters": sum(len(p) for p in prose),
        "native_tables": sum(1 for n in root.iter(W + "tbl") if not inside(n, W + "txbxContent")),
        "native_equations": sum(1 for _ in root.iter(M + "oMath")),
        "native_mark_allocations": len(marks),
        "mark_labels": dict(sorted(Counter(marks).items())),
        "underscore_paragraphs": len(answer_lines),
        "underscore_lengths": dict(sorted(Counter(len(p) for p in answer_lines).items())),
        "alignment_counts": dict(sorted(Counter(a for _, a, _ in prose_aligned).items())),
        "alignment_by_style": {s: dict(sorted(c.items())) for s, c in sorted(by_style.items())},
        "alignment_outliers_first_30": outliers[:30],
        "underscore_paragraphs_justified": justified_lines,
        "embedded_media": media,
        "drawing_objects": sum(counts.values()),
        "drawing_breakdown": dict(sorted(counts.items())),
        "image_placements_in_main_story": pictures,
        "picture_text_strips": strips,
        "vml_objects_without_drawingml": sum(1 for n in root.iter()
                                             if n.tag in (V + "line", V + "rect", V + "shape", V + "group",
                                                          V + "oval", V + "polyline", V + "curve", V + "arc")),
        "floating_text_boxes": len(floating),
        "floating_anchors_first_30": floating_anchors[:30],
        "layout_lints": lints,
        "header_footer_note_stories": sum(other_stories.values()),
        "tracked_revision_elements": revisions,
        "compatibility_mode": compat_mode,
        "main_story_text_sha256": hashlib.sha256(full_text.encode("utf-8")).hexdigest(),
        "prose_sha256": hashlib.sha256(text.encode("utf-8")).hexdigest(),
        "warnings": warnings,
        "note": "Counts do not prove editability of each image, line wrapping, pagination, or visual fidelity. Render and inspect every page.",
    }
    internal = {"text": text, "prose": [" ".join(p.split()) for p in prose],
                "aligned_prose": [(" ".join(t.split()), a) for t, a, _ in prose_aligned],
                "underscores": answer_lines, "marks": Counter(marks),
                "floating": Counter(map(compact, floating)), "stories": other_stories,
                "full_text": full_text, "equations": equations}
    summary["vml_objects"] = summary["vml_objects_without_drawingml"]  # Legacy key, now without fallback double counting.
    return summary, internal


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--baseline", type=Path,
                        help="Latest user-edited DOCX; compare text and editable line/mark content")
    parser.add_argument("--report", type=Path, help="New JSON report path; existing files are never overwritten")
    parser.add_argument("--strict", action="store_true", help="Return 3 when warnings exist; heuristics need review and are not automatic defects")
    args = parser.parse_args()
    try:
        report, current = read_docx(args.docx)
        if args.baseline:
            _, baseline = read_docx(args.baseline)
            report["baseline"] = str(args.baseline.resolve())
            report["preservation"] = {
                "main_story_prose_unchanged_ignoring_whitespace": current["text"] == baseline["text"],
                "main_story_text_unchanged_ignoring_whitespace": current["full_text"] == baseline["full_text"],
                "equation_structure_unchanged": current["equations"] == baseline["equations"],
                "underscore_paragraph_count_unchanged": len(current["underscores"]) == len(baseline["underscores"]),
                "underscore_line_count_unchanged": len(current["underscores"]) == len(baseline["underscores"]),
                "underscore_sequences_unchanged": current["underscores"] == baseline["underscores"],
                "mark_labels_unchanged": current["marks"] == baseline["marks"],
                "floating_text_unchanged_ignoring_position_and_whitespace": current["floating"] == baseline["floating"],
                "header_footer_note_text_unchanged": current["stories"] == baseline["stories"],
            }
            # Alignment the user set (e.g. justification) must survive layout work on unchanged paragraphs.
            base_items, cur_items = baseline["aligned_prose"], current["aligned_prose"]
            matcher = difflib.SequenceMatcher(None, [t for t, _ in base_items], [t for t, _ in cur_items],
                                              autojunk=False)
            changed = [{"text": base_items[b.a + k][0][:60], "baseline": base_items[b.a + k][1],
                        "revised": cur_items[b.b + k][1]}
                       for b in matcher.get_matching_blocks() for k in range(b.size)
                       if base_items[b.a + k][1] != cur_items[b.b + k][1]]
            matched = sum(b.size for b in matcher.get_matching_blocks())
            report["alignment_comparison_coverage"] = {"matched_paragraphs": matched,
                 "baseline_paragraphs": len(base_items), "revised_paragraphs": len(cur_items)}
            report["preservation"]["paragraph_alignment_unchanged_for_matching_paragraphs"] = not changed
            if changed:
                report["alignment_changes_first_30"] = changed[:30]
            if current["text"] != baseline["text"]:
                diff = difflib.unified_diff(baseline["prose"], current["prose"], "baseline", "revised",
                                            n=0, lineterm="")
                report["prose_diff_first_60_lines"] = list(diff)[:60]
        if args.report:
            with args.report.open("x", encoding="utf-8") as stream:
                stream.write(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
        # ASCII-safe stdout: Windows pipes use the ANSI code page and crash on CJK paths or NBSP.
        print(json.dumps(report, ensure_ascii=True, indent=2))
        if not all(report.get("preservation", {}).values()):
            return 2
        return 3 if args.strict and report["warnings"] else 0
    except (OSError, zipfile.BadZipFile, KeyError, ValueError, ET.ParseError) as error:
        print("Audit failed: " + json.dumps(str(error), ensure_ascii=True), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
