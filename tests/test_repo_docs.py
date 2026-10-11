import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


class RepoDocs(unittest.TestCase):
    def test_agents_md_has_hard_rules(self):
        text = (ROOT / 'AGENTS.md').read_text(encoding='utf-8')
        for must in ('Never commit the .xlsm', 'python ops/setup_dev.py', '.planning/STATE.md',
                     'python -m unittest discover -s tests -v', 'SPEC.md', 'No em-dashes'):
            self.assertIn(must, text)

    def test_claude_md_imports_agents(self):
        self.assertIn('@AGENTS.md', (ROOT / 'CLAUDE.md').read_text(encoding='utf-8'))

    def test_state_md_sections(self):
        text = (ROOT / '.planning' / 'STATE.md').read_text(encoding='utf-8')
        for heading in ('## Current phase', '## Last handoff', '## Next task'):
            self.assertIn(heading, text)

    def test_no_em_dashes_in_docs(self):
        for path in ('AGENTS.md', 'CLAUDE.md', '.planning/STATE.md', 'docs/DEPLOY.md'):
            self.assertNotIn('\u2014', (ROOT / path).read_text(encoding='utf-8'), path)
