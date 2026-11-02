"""校验发布表结构及直接 XML 导入的保真性。"""
from pathlib import Path
import importlib.util
import tempfile
import unittest
import zipfile
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
BOOK = ROOT/'原神角色养成管理表_v1.0.xlsx'

class WorkbookTests(unittest.TestCase):
    def test_blank_workbook_contract(self):
        w = load_workbook(BOOK)
        self.assertEqual(w.sheetnames, ['养成总表','今日练谁','圣遗物明细','角色参考卡','配装数据库','仪表盘','数据源','数据与更新','使用说明'])
        names = [w['养成总表'].cell(r,2).value for r in range(3,136)]
        self.assertEqual(len(set(names)),133)
        self.assertNotIn(None,names)
        self.assertEqual({w['养成总表'].cell(r,43).value for r in range(3,136)}, {'未拥有'})
        self.assertEqual(sum(bool(w['圣遗物明细'].cell(r,1).value) for r in range(3,668)),665)
        self.assertTrue(all(s.protection.sheet and not s.protection.password for s in w))
        self.assertFalse(w['养成总表']['D3'].protection.locked)
        self.assertTrue(w['养成总表']['AJ3'].protection.locked)
        for r in (7,13,19,25,31):
            self.assertFalse(w['今日练谁'][f'C{r}'].protection.locked)
        self.assertTrue(w.calculation.fullCalcOnLoad)
        self.assertFalse([(s.title,c.coordinate) for s in w for row in s for c in row if c.data_type=='e'])
        w.close()

    def test_xml_import_preserves_charts_and_formulas(self):
        spec=importlib.util.spec_from_file_location('miyoushe',ROOT/'_source/import_miyoushe.py')
        m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
        with tempfile.TemporaryDirectory() as td:
            dst=Path(td)/'synthetic.xlsx'
            m.edit_cells(str(BOOK),str(dst),{'养成总表':{'D3':80,'AQ3':'已有','AN3':'合成测试 & <安全>'}})
            with zipfile.ZipFile(BOOK) as a, zipfile.ZipFile(dst) as b:
                for name in a.namelist():
                    if name.startswith(('xl/charts/','xl/drawings/')) or name=='xl/styles.xml':
                        self.assertEqual(a.read(name),b.read(name),name)
            w=load_workbook(dst)
            self.assertEqual(w['养成总表']['D3'].value,80)
            self.assertEqual(w['养成总表']['AN3'].value,'合成测试 & <安全>')
            self.assertTrue(w['养成总表']['AJ3'].value.startswith('='))
            self.assertEqual(len(w['仪表盘']._charts),2)
            w.close()

if __name__=='__main__': unittest.main()
