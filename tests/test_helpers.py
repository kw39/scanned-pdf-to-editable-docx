"""Synthetic regression cases; never load the user's source papers."""
import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

from PIL import Image, ImageDraw
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parents[1]


def load_helper(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / 'scripts' / (name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


audit = load_helper('audit_docx')
fit = load_helper('compare_line_endings')
scan = load_helper('measure_scan_vertical')
W = audit.W
M = audit.M


def paragraph(text, properties=None):
    p = ET.Element(W + 'p')
    if properties is not None:
        p.append(properties)
    r = ET.SubElement(p, W + 'r')
    ET.SubElement(r, W + 't').text = text
    return p


def properties(alignment=None, style=None, spacing=None):
    node = ET.Element(W + 'pPr')
    if alignment:
        ET.SubElement(node, W + 'jc', {W + 'val': alignment})
    if style:
        ET.SubElement(node, W + 'pStyle', {W + 'val': style})
    if spacing:
        ET.SubElement(node, W + 'spacing', {W + k: str(v) for k, v in spacing.items()})
    return node


def equation(p, swapped=False):
    math = ET.SubElement(p, M + 'oMath')
    frac = ET.SubElement(math, M + 'f')
    for tag, value in [('num', 'a'), ('den', 'b')]:
        element = ET.SubElement(frac, M + tag)
        run = ET.SubElement(element, M + 'r')
        ET.SubElement(run, M + 't').text = ('b' if value == 'a' else 'a') if swapped else value
    return p


def make_docx(path, paragraphs=None, styles=None, footer=None):
    document = ET.Element(W + 'document')
    body = ET.SubElement(document, W + 'body')
    for p in paragraphs if paragraphs is not None else [paragraph('First question.'), paragraph('Last question.')]:
        body.append(p)
    section = ET.SubElement(body, W + 'sectPr')
    ET.SubElement(section, W + 'pgSz', {W + 'w': '12240', W + 'h': '15840'})
    settings = ET.Element(W + 'settings')
    compat = ET.SubElement(settings, W + 'compat')
    ET.SubElement(compat, W + 'compatSetting', {W + 'name': 'compatibilityMode', W + 'val': '15'})
    if styles is None:
        styles = ET.Element(W + 'styles')
    # Minimal packages for structural tests; desktop rendering is tested separately.
    with zipfile.ZipFile(path, 'w') as package:
        package.writestr('word/document.xml', ET.tostring(document))
        package.writestr('word/settings.xml', ET.tostring(settings))
        package.writestr('word/styles.xml', ET.tostring(styles))
        if footer is not None:
            root = ET.Element(W + 'ftr')
            root.append(paragraph(footer))
            package.writestr('word/footer1.xml', ET.tostring(root))
    return path


def style_tree(defaults=None, derived=True):
    root = ET.Element(W + 'styles')
    if defaults:
        doc = ET.SubElement(root, W + 'docDefaults')
        pp = ET.SubElement(ET.SubElement(doc, W + 'pPrDefault'), W + 'pPr')
        ET.SubElement(pp, W + 'spacing', {W + k: str(v) for k, v in defaults.items()})
    base = ET.SubElement(root, W + 'style', {W + 'type': 'paragraph', W + 'styleId': 'Base', W + 'default': '1'})
    pp = ET.SubElement(base, W + 'pPr')
    ET.SubElement(pp, W + 'jc', {W + 'val': 'both'})
    ET.SubElement(pp, W + 'spacing', {W + 'lineRule': 'exact', W + 'line': '300'})
    if derived:
        child = ET.SubElement(root, W + 'style', {W + 'type': 'paragraph', W + 'styleId': 'Question'})
        ET.SubElement(child, W + 'basedOn', {W + 'val': 'Base'})
    return root


def make_pdf(path, rows=None, extra_page=None):
    pdf = canvas.Canvas(str(path), pagesize=(612, 792))
    pdf.setFont('Helvetica', 11)
    for text, top in rows if rows is not None else [
        ('1. Explain how supply changes', 100),
        ('when the price of a good rises.', 114),
        ('2. State one limitation. (2 marks)', 138),
        ('____________________', 170), ('____________________', 188.5),
        ('P.1 of 1', 755),
    ]:
        pdf.drawString(60, 792 - top, text)
    if extra_page is not None:
        pdf.showPage()
        pdf.setFont('Helvetica', 11)
        for text, top in extra_page:
            pdf.drawString(60, 792 - top, text)
    pdf.save()
    return path


class TemporaryCase(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.dir = Path(self.temp.name)

    def cli(self, helper, *args):
        return subprocess.run([sys.executable, str(ROOT / 'scripts' / helper), *map(str, args)],
                              capture_output=True, text=True, encoding='utf-8', timeout=30)


class AuditTests(TemporaryCase):
    def test_prose_hash_covers_every_paragraph_and_excludes_answer_lines(self):
        path = make_docx(self.dir / 'a.docx', [paragraph('First question.'), paragraph('Last question.'), paragraph('__________')])
        report, internal = audit.read_docx(path)
        expected = 'Firstquestion.Lastquestion.'
        self.assertEqual(internal['text'], expected)
        self.assertEqual(report['prose_sha256'], hashlib.sha256(expected.encode()).hexdigest())
        self.assertEqual(report['underscore_paragraphs'], 1)

    def test_earlier_edit_is_detected_when_last_paragraph_is_unchanged(self):
        base = make_docx(self.dir / 'base.docx')
        new = make_docx(self.dir / 'new.docx', [paragraph('Changed question.'), paragraph('Last question.')])
        result = self.cli('audit_docx.py', new, '--baseline', base)
        self.assertEqual(result.returncode, 2, result.stderr)
        report = json.loads(result.stdout)
        self.assertFalse(report['preservation']['main_story_prose_unchanged_ignoring_whitespace'])
        self.assertTrue(report['prose_diff_first_60_lines'])

    def test_identical_baseline_passes(self):
        path = make_docx(self.dir / 'a.docx')
        result = self.cli('audit_docx.py', path, '--baseline', path)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue(all(json.loads(result.stdout)['preservation'].values()))

    def test_inherited_justify_and_exact_equation_risk_are_resolved(self):
        p = equation(paragraph('A fraction: ', properties(style='Question')))
        report, _ = audit.read_docx(make_docx(self.dir / 'a.docx', [p], style_tree()))
        self.assertEqual(report['alignment_counts'], {'justify': 1})
        risk = report['layout_lints']['equation_clipping_risks'][0]
        self.assertEqual(risk['rule_source'], 'style:Base')
        self.assertEqual(risk['line_twips'], '300')

    def test_direct_at_least_override_clears_exact_spacing_warning(self):
        p = equation(paragraph('A fraction: ', properties(style='Question', spacing={'lineRule': 'atLeast', 'line': 300})))
        report, _ = audit.read_docx(make_docx(self.dir / 'a.docx', [p], style_tree()))
        self.assertEqual(report['layout_lints']['tall_equations_with_exact_spacing'], 0)

    def test_document_default_spacing_is_checked(self):
        styles = ET.Element(W + 'styles')
        pp = ET.SubElement(ET.SubElement(ET.SubElement(styles, W + 'docDefaults'), W + 'pPrDefault'), W + 'pPr')
        ET.SubElement(pp, W + 'spacing', {W + 'lineRule': 'exact', W + 'line': '200'})
        report, _ = audit.read_docx(make_docx(self.dir / 'a.docx', [equation(paragraph('Fraction: '))], styles))
        self.assertEqual(report['layout_lints']['equation_clipping_risks'][0]['rule_source'], 'docDefaults')

    def test_alignment_change_is_reported(self):
        base = make_docx(self.dir / 'base.docx', [paragraph('Same words.', properties(alignment='both'))])
        new = make_docx(self.dir / 'new.docx', [paragraph('Same words.', properties(alignment='left'))])
        result = self.cli('audit_docx.py', new, '--baseline', base)
        self.assertEqual(result.returncode, 2, result.stderr)
        report = json.loads(result.stdout)
        self.assertFalse(report['preservation']['paragraph_alignment_unchanged_for_matching_paragraphs'])

    def test_equation_structure_change_is_reported(self):
        base = make_docx(self.dir / 'base.docx', [equation(paragraph('Fraction: '))])
        new = make_docx(self.dir / 'new.docx', [equation(paragraph('Fraction: '), swapped=True)])
        result = self.cli('audit_docx.py', new, '--baseline', base)
        self.assertEqual(result.returncode, 2, result.stderr)
        self.assertFalse(json.loads(result.stdout)['preservation']['equation_structure_unchanged'])

    def test_footer_change_is_reported(self):
        base = make_docx(self.dir / 'base.docx', footer='P.1 of 3')
        new = make_docx(self.dir / 'new.docx', footer='P.1 of 4')
        result = self.cli('audit_docx.py', new, '--baseline', base)
        self.assertEqual(result.returncode, 2, result.stderr)
        self.assertFalse(json.loads(result.stdout)['preservation']['header_footer_note_text_unchanged'])

    def test_mark_and_underscore_changes_are_reported(self):
        base = make_docx(self.dir / 'base.docx', [paragraph('Question. (6 marks)'), paragraph('__________')])
        new = make_docx(self.dir / 'new.docx', [paragraph('Question. (4 marks)'), paragraph('____________')])
        result = self.cli('audit_docx.py', new, '--baseline', base)
        self.assertEqual(result.returncode, 2, result.stderr)
        preservation = json.loads(result.stdout)['preservation']
        self.assertFalse(preservation['mark_labels_unchanged'])
        self.assertFalse(preservation['underscore_sequences_unchanged'])

    def test_no_native_prose_produces_strict_warning(self):
        path = make_docx(self.dir / 'empty.docx', [])
        result = self.cli('audit_docx.py', path, '--strict')
        self.assertEqual(result.returncode, 3, result.stderr)
        self.assertTrue(json.loads(result.stdout)['warnings'])

    def test_report_cannot_replace_input_or_existing_output(self):
        path = make_docx(self.dir / 'a.docx')
        before = path.read_bytes()
        result = self.cli('audit_docx.py', path, '--report', path)
        self.assertEqual(result.returncode, 1)
        self.assertEqual(path.read_bytes(), before)
        output = self.dir / 'audit.json'
        output.write_text('keep me', encoding='utf-8')
        self.assertEqual(self.cli('audit_docx.py', path, '--report', output).returncode, 1)
        self.assertEqual(output.read_text(), 'keep me')

    def test_new_utf8_report_is_written(self):
        path = make_docx(self.dir / 'a.docx', [paragraph('中文題目。')])
        output = self.dir / 'audit.json'
        result = self.cli('audit_docx.py', path, '--report', output)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(output.read_text(encoding='utf-8'))['native_prose_characters'], 5)

    def test_invalid_docx_fails_without_traceback(self):
        path = self.dir / 'bad.docx'
        path.write_text('not a zip', encoding='utf-8')
        result = self.cli('audit_docx.py', path)
        self.assertEqual(result.returncode, 1)
        self.assertNotIn('Traceback', result.stderr)


class ComparisonTests(TemporaryCase):
    def setUp(self):
        super().setUp()
        self.pdf = make_pdf(self.dir / 'render.pdf')
        self.manifest = json.loads((ROOT / 'examples/manifest.json').read_text(encoding='utf-8'))

    def compare(self, manifest=None, pdf=None):
        path = self.dir / 'manifest.json'
        path.write_text(json.dumps(self.manifest if manifest is None else manifest), encoding='utf-8')
        return self.cli('compare_line_endings.py', pdf or self.pdf, '--manifest', path)

    def test_complete_example_passes_with_spacing_and_answer_lines(self):
        result = self.compare()
        self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
        report = json.loads(result.stdout)
        self.assertTrue(report['all_checks_passed'])
        self.assertTrue(report['vertical_spacing_preserved'])
        self.assertEqual(report['answer_lines'][0]['rendered_count'], 2)

    def test_same_line_count_with_different_endings_is_accepted(self):
        block = {'id': 'q', 'page': 1, 'lines': ['Alpha beta', 'gamma delta']}
        result = fit.compare_block(block, ['Alpha', 'beta gamma delta'])
        self.assertTrue(result['line_count_matches'])
        self.assertFalse(result['exact_line_endings'])

    def test_extra_line_is_a_failure(self):
        rows = [('1. Explain how',100),('supply changes',114),('when the price of a good rises.',128),
                ('2. State one limitation. (2 marks)',160)]
        pdf = make_pdf(self.dir / 'wrapped.pdf', rows)
        manifest = copy.deepcopy(self.manifest)
        for block in manifest['pages'][0]['blocks']:
            block.pop('render_bbox', None)
            block.pop('spacing', None)
        manifest['pages'][0].pop('answer_line_count')
        manifest['pages'][0].pop('answer_line_pitch_pt')
        result = self.compare(manifest, pdf)
        self.assertEqual(result.returncode, 2, result.stderr)
        self.assertEqual(json.loads(result.stdout)['line_count_mismatches'][0]['render_line_count'], 3)

    def test_text_on_wrong_page_stays_unresolved(self):
        pdf = make_pdf(self.dir / 'wrong-page.pdf', [('Other text', 100)], [('1. Explain how supply changes',100),('when the price of a good rises.',114)])
        result = self.compare(pdf=pdf)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(json.loads(result.stdout)['unresolved_blocks'][0]['found_on_other_render_pages'], [2])

    def test_empty_reference_cannot_pass(self):
        result = self.compare({'reference_verified': True, 'blocks': []})
        self.assertEqual(result.returncode, 3, result.stderr)
        self.assertIsNone(json.loads(result.stdout)['all_checks_passed'])

    def test_unverified_reference_cannot_pass(self):
        self.manifest['reference_verified'] = False
        result = self.compare()
        self.assertEqual(result.returncode, 3, result.stderr)
        self.assertIsNone(json.loads(result.stdout)['all_checks_passed'])

    def test_missing_requested_next_block_is_incomplete(self):
        self.manifest['pages'][0]['blocks'] = self.manifest['pages'][0]['blocks'][:1]
        result = self.compare()
        self.assertEqual(result.returncode, 3, result.stderr)
        report = json.loads(result.stdout)
        self.assertIsNone(report['blocks'][0]['vertical']['baseline_to_next_block']['ok'])

    def test_uneven_pitch_cannot_hide_behind_a_matching_median(self):
        result = fit.check_pitch(15, [10, 15, 20], 1)
        self.assertEqual(result['actual'], 15)
        self.assertFalse(result['ok'])

    def test_undeclared_discretionary_hyphen_stays_strict(self):
        block = {'id': 'q', 'page': 1, 'lines': ['Explain consump-', 'tion of goods.']}
        self.assertEqual(fit.compare_block(block, ['Explain consumption', 'of goods.'])['status'], 'unmatched_text')
        block['discretionary_hyphen_breaks'] = [1]
        result = fit.compare_block(block, ['Explain consumption', 'of goods.'])
        self.assertTrue(result['line_count_matches'])
        self.assertEqual(result['hyphenation_variant'], 'joined_word')

    def test_duplicate_text_is_ambiguous(self):
        result = fit.compare_block({'id': 'q', 'page': 1, 'lines': ['Same text']}, ['Same text', 'Same text'])
        self.assertEqual(result['status'], 'ambiguous_text')
        self.assertIsNone(result['line_count_matches'])

    def test_duplicate_ids_and_bad_bbox_are_rejected(self):
        self.manifest['pages'][0]['blocks'][1]['id'] = 'q1'
        self.assertEqual(self.compare().returncode, 1)
        self.manifest['pages'][0]['blocks'][1]['id'] = 'q2'
        self.manifest['pages'][0]['blocks'][0]['render_bbox'] = [60, 100, 60, 130]
        self.assertEqual(self.compare().returncode, 1)

    def test_cjk_and_native_superscripts_remain_tokens(self):
        self.assertEqual(fit.tokens('供應 x² = 10³'), ['供', '應', 'x2', '=', '103'])

    def test_subscript_glyphs_join_the_same_pdf_row(self):
        class Page:
            chars = [
                {'text':'S','top':10,'bottom':20,'x0':0,'x1':7,'size':10},
                {'text':'1','top':17,'bottom':23,'x0':7,'x1':10,'size':6},
                {'text':'=','top':10,'bottom':20,'x0':15,'x1':22,'size':10},
            ]
        rows = fit.pdf_rows(Page())
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][0], 'S1 =')

    def test_json_report_cannot_replace_render_or_manifest(self):
        path = self.dir / 'manifest.json'
        path.write_text(json.dumps(self.manifest), encoding='utf-8')
        for target in (path, self.pdf):
            before = target.read_bytes()
            result = self.cli('compare_line_endings.py', self.pdf, '--manifest', path, '--json', target)
            self.assertEqual(result.returncode, 1, result.stderr)
            self.assertEqual(target.read_bytes(), before)

    def test_single_line_pitch_request_is_incomplete(self):
        self.manifest['pages'][0]['blocks'][1]['spacing'] = {'observed_baseline_pitch_pt': 14}
        result = self.compare()
        self.assertEqual(result.returncode, 3, result.stderr)
        self.assertIsNone(json.loads(result.stdout)['blocks'][1]['vertical']['baseline_pitch']['ok'])

    def test_plain_text_requires_explicit_boundaries_and_verification(self):
        path = self.dir / 'source.txt'
        path.write_text('#page 1\n#block q1\n1. Explain how supply changes\nwhen the price of a good rises.\n', encoding='utf-8')
        result = self.cli('compare_line_endings.py', self.pdf, '--scan-lines', path)
        self.assertEqual(result.returncode, 3, result.stderr)
        result = self.cli('compare_line_endings.py', self.pdf, '--scan-lines', path, '--verified-reference')
        self.assertEqual(result.returncode, 0, result.stderr)


class ScanTests(TemporaryCase):
    def fixture(self):
        image = Image.new('L', (200, 160), 255)
        draw = ImageDraw.Draw(image)
        for y in (10, 40, 70):
            for x in (10, 40, 70):
                draw.rectangle((x, y, x+15, y+7), fill=0)
        for y in (110, 140):
            draw.line((0,y,199,y), fill=0, width=1)
        return image

    def test_known_text_and_rule_pitch_are_measured(self):
        report = scan.measure(self.fixture(), None, 150, 0.5)
        self.assertEqual(len(report['text_lines']), 3)
        self.assertEqual(report['text_line_steps_pt'], [15,15])
        self.assertEqual(report['ruled_lines'], 2)
        self.assertEqual(report['ruled_line_median_pitch_pt'], 15)

    def test_invalid_scale_bbox_and_threshold_are_rejected(self):
        for scale in (0, -1, float('nan'), float('inf')):
            with self.assertRaises(ValueError):
                scan.measure(self.fixture(), None, 150, scale)
        with self.assertRaises(ValueError):
            scan.measure(self.fixture(), [0,0,201,160], 150, 1)
        with self.assertRaises(ValueError):
            scan.measure(self.fixture(), None, 256, 1)

    def test_image_cli_measurement_and_missing_scale_are_clear(self):
        path = self.dir / 'scan.png'
        self.fixture().save(path)
        result = self.cli('measure_scan_vertical.py', path, '--column-pt', 100, '--column-px', 200)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)['pt_per_px'], 0.5)
        result = self.cli('measure_scan_vertical.py', path, '--column-pt', 100)
        self.assertEqual(result.returncode, 1)
        self.assertNotIn('Traceback', result.stderr)


if __name__ == '__main__':
    unittest.main()
