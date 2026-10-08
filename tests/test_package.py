"""Validate discoverability, reference resolution and public packaging."""
import json
from pathlib import Path
import re
import unittest

import yaml

ROOT = Path(__file__).resolve().parents[1]


class PackageTests(unittest.TestCase):
    def test_skill_frontmatter_has_a_discriminating_description_and_license(self):
        source = (ROOT / 'SKILL.md').read_text(encoding='utf-8')
        match = re.match(r'^---\n(.*?)\n---', source, re.S)
        self.assertIsNotNone(match)
        metadata = yaml.safe_load(match.group(1))
        self.assertEqual(metadata['name'], ROOT.name)
        self.assertRegex(metadata['name'], r'^[a-z0-9]+(?:-[a-z0-9]+)*$')
        self.assertLessEqual(len(metadata['name']), 64)
        self.assertTrue(30 < len(metadata['description']) <= 1024)
        self.assertEqual(metadata['license'], 'MIT')

    def test_ui_metadata_matches_skill_and_preserves_automatic_selection(self):
        metadata = yaml.safe_load((ROOT / 'agents/openai.yaml').read_text(encoding='utf-8'))
        interface = metadata['interface']
        self.assertTrue(25 <= len(interface['short_description']) <= 64)
        self.assertIn('$scanned-pdf-to-editable-docx', interface['default_prompt'])
        self.assertTrue(metadata.get('policy', {}).get('allow_implicit_invocation', True))

    def test_local_markdown_links_and_anchors_resolve(self):
        for document in [ROOT / 'README.md', ROOT / 'SKILL.md', *sorted((ROOT / 'references').glob('*.md'))]:
            for link in re.findall(r'\[[^\]]+\]\(([^)]+)\)', document.read_text(encoding='utf-8')):
                if re.match(r'[a-z]+://', link):
                    continue
                target, _, anchor = link.partition('#')
                destination = (document.parent / target).resolve() if target else document
                self.assertTrue(destination.is_relative_to(ROOT.resolve()), f'{document}: {link}')
                self.assertTrue(destination.is_file(), f'{document}: {link}')
                if anchor:
                    headings = re.findall(r'^#+\s+(.+)$', destination.read_text(encoding='utf-8'), re.M)
                    slugs = [re.sub(r'[^\w -]', '', heading.lower()).replace(' ', '-') for heading in headings]
                    self.assertIn(anchor, slugs, f'{document}: {link}')

    def test_example_manifest_is_verified_and_has_unique_source_blocks(self):
        manifest = json.loads((ROOT / 'examples/manifest.json').read_text(encoding='utf-8'))
        self.assertIs(manifest['reference_verified'], True)
        self.assertEqual([p['physical_page'] for p in manifest['pages']], list(range(1, len(manifest['pages'])+1)))
        blocks = [b for p in manifest['pages'] for b in p['blocks']]
        self.assertTrue(blocks)
        self.assertEqual(len({b['id'] for b in blocks}), len(blocks))
        for block in blocks:
            self.assertTrue(block['lines'])

    def test_ci_only_requests_read_permission_and_runs_regressions(self):
        workflow = yaml.load((ROOT / '.github/workflows/validate.yml').read_text(encoding='utf-8'), Loader=yaml.BaseLoader)
        self.assertEqual(workflow['permissions'], {'contents': 'read'})
        self.assertIn('pull_request', workflow['on'])
        job = workflow['jobs']['helpers']
        self.assertEqual(set(job['strategy']['matrix']['python-version']), {'3.10', '3.12'})
        self.assertEqual(set(job['strategy']['matrix']['os']), {'ubuntu-latest', 'windows-latest'})
        self.assertTrue(any('unittest discover' in step.get('run', '') for step in job['steps']))

    def test_release_sources_do_not_contain_private_material(self):
        for folder in ['scripts', 'references', 'agents', 'examples', 'tests']:
            for path in (ROOT / folder).rglob('*'):
                if not path.is_file() or '__pycache__' in path.parts:
                    continue
                self.assertNotIn(path.suffix.lower(), {'.docx','.pdf','.png','.jpg','.zip','.pyc'})
                source = path.read_text(encoding='utf-8')
                self.assertIsNone(re.search(r'[A-Za-z]:[\\/]Users[\\/]|[A-Za-z]:[\\/]gpt[\\/]', source), str(path))


if __name__ == '__main__':
    unittest.main()
