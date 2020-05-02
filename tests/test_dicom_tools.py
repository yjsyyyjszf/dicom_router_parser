import unittest
import os
import re
import pathlib
from pyapp.pylibs.dicom_tools import *

BASE_DIR, SCRIPT_NAME = os.path.split(os.path.abspath(__file__))
PARENT_PATH, CURR_DIR = os.path.split(BASE_DIR)


class TestDicomTools(unittest.TestCase):
    def setUp(self):
        self.valid_dir = pathlib.Path(PARENT_PATH, 'input', 'DICOM')
        self.valid_dcmtk = pathlib.Path(PARENT_PATH, 'input', 'tag_dumps',
                                        'dcmtk_dump_src.txt')
        self.valid_fuji = pathlib.Path(PARENT_PATH, 'input', 'tag_dumps',
                                       'fuji_dicom_dump.txt')

    def test_build_dcmtk_tag_dict(self):
        dcmtk_tag_dict = build_dcmtk_tag_dict(self.valid_dcmtk)
        # match: (0008,0050)
        for key, tag in dcmtk_tag_dict.items():
            if 'filename' not in key:
                match = re.search("^(\\([0-9]{4},)([0-9]{4}\\))$", tag)
                self.assertIsNotNone(match)
                self.assertIsInstance(tag, str)
                self.assertEqual(tag[match.start():match.end()], tag)

    def test_build_fuji_tag_dict(self):
        fuji_tag_dict = build_fuji_tag_dict(self.valid_fuji)
        # match: 0008 0050
        for key, tag in fuji_tag_dict.items():
            if 'filename' not in key:
                match = re.search("^[0-9]{4} [0-9]{4}$", tag)
                self.assertIsNotNone(match)
                self.assertIsInstance(tag, str)
                self.assertEqual(tag[match.start():match.end()], tag)

    def tearDown(self) -> None:
        pass

if __name__ == '__main__':
    unittest.main()
