"""Installer/app integration without changing the machine's installed tools."""
import importlib.util
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('sova_platform', ROOT / 'app/platform_utils.py')
platform = importlib.util.module_from_spec(spec)
spec.loader.exec_module(platform)


class WindowsRuntimeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        (self.root / 'app').mkdir()
        self.env = patch.dict(os.environ, {'PATH': 'original-path'}, clear=False)
        self.env.start()
        self.addCleanup(self.env.stop)
        self.windows = patch.object(platform, 'IS_WIN', True)
        self.base = patch.object(platform, 'BASE', self.root / 'app')
        self.windows.start(); self.base.start()
        self.addCleanup(self.windows.stop); self.addCleanup(self.base.stop)

    def write_config(self, value):
        (self.root / 'runtime-tools.json').write_text(json.dumps(value), encoding='utf-8')

    def test_private_tool_paths_and_russian_model(self):
        tools = self.root / 'tools'; tools.mkdir()
        model = self.root / 'tessdata'; model.mkdir()
        (model / 'rus.traineddata').write_bytes(b'model')
        self.write_config({'tool_dirs': [str(tools), str(self.root / 'missing')],
                           'tessdata_dir': str(model)})
        platform.configure_windows_tools()
        self.assertEqual(os.environ['PATH'].split(os.pathsep), [str(tools), 'original-path'])
        self.assertEqual(os.environ['TESSDATA_PREFIX'], str(model))

    def test_broken_config_does_not_break_startup(self):
        (self.root / 'runtime-tools.json').write_text('{bad json')
        self.assertEqual(platform.configure_windows_tools(), {})
        self.assertEqual(os.environ['PATH'], 'original-path')
        self.write_config([])
        self.assertEqual(platform.configure_windows_tools(), {})
        self.write_config({'tool_dirs': 'not-a-list'})
        self.assertEqual(platform.configure_windows_tools(), {})

    def test_detected_libreoffice_location_is_used(self):
        office = self.root / 'soffice.exe'; office.touch()
        with patch.object(platform, 'WINDOWS_RUNTIME', {'soffice': str(office)}):
            self.assertEqual(platform.soffice_cmd(), str(office))


if __name__ == '__main__':
    unittest.main()
