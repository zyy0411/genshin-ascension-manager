"""离线发布回归检查，不访问真实账号。"""
from pathlib import Path
import hashlib
import subprocess
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]

class ReleaseTests(unittest.TestCase):
    def test_assemble_from_root(self):
        result = subprocess.run([sys.executable, '_source/assemble_v4.py'], cwd=ROOT, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        parts = ['v4_common.py', 'v4_build_a.py', 'v4_build_b.py', 'v4_build_c.py', 'v4_build_d.py']
        self.assertEqual((ROOT/'_source/build_tracker_v4.py').read_text(), '\n\n'.join((ROOT/'_source'/p).read_text() for p in parts))

    def test_offline_rebuild_uses_shipped_data(self):
        paths = sorted((ROOT/'_source').glob('wiki_*.json'))
        before = [hashlib.sha256(p.read_bytes()).hexdigest() for p in paths]
        result = subprocess.run([sys.executable, '_source/update_data.py', '--offline'], cwd=ROOT, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(before, [hashlib.sha256(p.read_bytes()).hexdigest() for p in paths])

    def test_reference_overrides_are_independent(self):
        from openpyxl import load_workbook
        w = load_workbook(ROOT/'原神角色养成管理表_v1.0.xlsx', read_only=True)
        for col in 'BCDEFGH':
            self.assertTrue(w['角色参考卡'][col+'31'].value.startswith('=IF('+col+'32<>"",'+col+'32,'), col)
        w.close()

    def test_missing_cookie_is_readable(self):
        result = subprocess.run([sys.executable, '_source/import_miyoushe.py', '--check'], cwd=ROOT, capture_output=True, text=True)
        self.assertEqual(result.returncode, 1)
        self.assertIn('Cookie', result.stdout)
        self.assertNotIn('Traceback', result.stderr)
        self.assertFalse(list((ROOT/'_wiki_tmp').rglob('raw_*.json')))

if __name__ == '__main__':
    unittest.main()
