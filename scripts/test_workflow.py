"""Entry-point regression tests. Python stdlib only; no Word or fonts required."""
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('gongwen_entry', ROOT / 'scripts/gongwen.py')
gongwen = importlib.util.module_from_spec(spec)
spec.loader.exec_module(gongwen)


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix='.gongwen-workflow-', dir=Path.cwd())
        self.work = Path(self.directory.name)
        self.source = self.work / 'source.md'
        self.source.write_text('# Test\n\nBody', encoding='utf-8')

    def tearDown(self):
        self.directory.cleanup()

    def test_prepare_does_not_publish_interrupted_conversion(self):
        target = self.work / 'output.docx'
        def failed_converter(args, **kwargs):
            Path(args[args.index('-o') + 1]).write_bytes(b'incomplete')
            raise RuntimeError('converter interrupted after writing output')
        with patch.object(gongwen, 'executable', return_value='pandoc'), patch.object(gongwen, 'run', side_effect=failed_converter):
            with self.assertRaises(RuntimeError):
                gongwen.prepare(self.source, target)
        self.assertFalse(target.exists())
        self.assertFalse(list(self.work.glob('.gongwen-prepare-*')))

    def test_prepare_refuses_file_created_during_conversion(self):
        target = self.work / 'output.docx'
        def raced_converter(args, **kwargs):
            Path(args[args.index('-o') + 1]).write_bytes(b'converted')
            target.write_bytes(b'another process owns this file')
            return ''
        with patch.object(gongwen, 'executable', return_value='pandoc'), patch.object(gongwen, 'run', side_effect=raced_converter):
            with self.assertRaises(FileExistsError):
                gongwen.prepare(self.source, target)
        self.assertEqual(target.read_bytes(), b'another process owns this file')
        self.assertFalse(list(self.work.glob('.gongwen-prepare-*')))

    def test_prepare_publishes_complete_docx_copy(self):
        source = self.work / 'source.docx'
        source.write_bytes(b'complete source bytes')
        target = self.work / 'output.docx'
        gongwen.prepare(source, target)
        self.assertEqual(source.read_bytes(), target.read_bytes())
        self.assertFalse(list(self.work.glob('.gongwen-prepare-*')))

    def invoke_check(self, output, engine_script):
        with patch.object(gongwen, 'engine', return_value=[sys.executable, engine_script]), patch.object(sys, 'argv', ['gongwen.py', 'check', str(self.source), '--output', str(output)]):
            gongwen.main()

    def test_failed_check_keeps_json_report_and_exit_status(self):
        fake_engine = self.work / 'engine.py'
        data = {'schemaErrors': [], 'formatErrors': ['bad spacing'], 'contentPreserved': True}
        fake_engine.write_text('import sys\nprint(' + repr(json.dumps(data)) + ')\nsys.exit(2)\n', encoding='utf-8')
        output = self.work / 'report.json'
        with self.assertRaises(SystemExit) as error:
            self.invoke_check(output, fake_engine)
        self.assertEqual(error.exception.code, 2)
        self.assertEqual(json.loads(output.read_text(encoding='utf-8')), data)

    def test_check_refuses_existing_report_before_engine_execution(self):
        output = self.work / 'report.json'
        output.write_text('existing report', encoding='utf-8')
        with patch.object(gongwen, 'engine', side_effect=AssertionError('engine must not run')), patch.object(sys, 'argv', ['gongwen.py', 'check', str(self.source), '--output', str(output)]):
            with self.assertRaises(ValueError):
                gongwen.main()
        self.assertEqual(output.read_text(encoding='utf-8'), 'existing report')

    def test_failed_engine_without_json_does_not_publish_report(self):
        fake_engine = self.work / 'engine.py'
        fake_engine.write_text("import sys\nprint('cannot open source',file=sys.stderr)\nsys.exit(2)\n", encoding='utf-8')
        output = self.work / 'report.json'
        with self.assertRaisesRegex(RuntimeError, 'cannot open source'):
            self.invoke_check(output, fake_engine)
        self.assertFalse(output.exists())

    def test_skill_archive_contains_source_and_excludes_caches(self):
        import zipfile
        output = self.work / 'skill.zip'
        result = subprocess.run([sys.executable, str(ROOT / 'scripts/pack_skill.py'), '--output', str(output)], capture_output=True, text=True, encoding='utf-8')
        self.assertEqual(result.returncode, 0, result.stderr)
        with zipfile.ZipFile(output) as archive:
            names = archive.namelist()
        self.assertIn(ROOT.name + '/SKILL.md', names)
        self.assertIn(ROOT.name + '/assets/profiles/gbt9704.json', names)
        self.assertIn(ROOT.name + '/examples/notice.md', names)
        self.assertIn(ROOT.name + '/LICENSE', names)
        self.assertNotIn(ROOT.name + '/assets/reference.docx', names)
        self.assertFalse(any(set(Path(name).parts) & {'bin', 'obj', '__pycache__', '.gongwen-cache', '.git'} for name in names))
        original = output.read_bytes()
        result = subprocess.run([sys.executable, str(ROOT / 'scripts/pack_skill.py'), '--output', str(output)], capture_output=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(output.read_bytes(), original)

    def test_skill_archive_excludes_work_files_and_its_own_output(self):
        import shutil
        import zipfile
        root = self.work / 'isolated-skill'
        (root / 'scripts').mkdir(parents=True)
        (root / 'dist').mkdir()
        (root / 'private-work').mkdir()
        (root / 'SKILL.md').write_text('skill', encoding='utf-8')
        (root / 'private-work/input.docx').write_bytes(b'private input')
        (root / 'local-report.json').write_text('{}', encoding='utf-8')
        script = root / 'scripts/pack_skill.py'
        shutil.copy2(ROOT / 'scripts/pack_skill.py', script)
        output = root / 'dist/skill.zip'
        subprocess.run([sys.executable, str(script), '--output', str(output)], check=True, capture_output=True)
        with zipfile.ZipFile(output) as archive:
            self.assertEqual(set(archive.namelist()), {
                'isolated-skill/SKILL.md', 'isolated-skill/scripts/pack_skill.py'})


if __name__ == '__main__':
    unittest.main()
