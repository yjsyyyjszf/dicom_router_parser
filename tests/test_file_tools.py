import unittest
import os
import pathlib
from sys import platform
import shutil

from pyapp.pylibs.file_tools import *

BASE_DIR, SCRIPT_NAME = os.path.split(os.path.abspath(__file__))
PARENT_PATH, CURR_DIR = os.path.split(BASE_DIR)


class TestFileTools(unittest.TestCase):
    """Test case class for /pyapp/pylibs/file_tools.py"""

    def setUp(self):
        self.valid_dir = pathlib.Path(PARENT_PATH, 'input', 'DICOM')
        self.valid_file = pathlib.Path(BASE_DIR, '__init__.py')
        self.noext_file = pathlib.Path(PARENT_PATH, 'LICENSE')
        self.valid_parent = pathlib.Path(PARENT_PATH, 'input')
        self.invalid_path = pathlib.Path(PARENT_PATH, 'does', 'not', 'exist')
        self.valid_ext = ['.txt', '.dcm']
        self.is_win = platform.startswith('win')
        self.out_path = os.path.join(BASE_DIR, 'output')

    def test_build_index_alphabet(self):
        alpha_dict = build_index_alphabet()
        self.assertIsInstance(alpha_dict, dict)
        self.assertEqual(alpha_dict[1], 'A')
        self.assertEqual(alpha_dict[26], 'Z')
        self.assertEqual(alpha_dict[27], 'AA')
        self.assertEqual(alpha_dict[53], 'BA')
        self.assertEqual(alpha_dict[79], 'CA')

    def test_build_extension_count_str(self):
        ext_str = build_extension_count_str(self.valid_dir)
        self.assertIsInstance(ext_str, str)
        self.assertTrue("files" in ext_str)
        ext_str = build_extension_count_str(self.valid_parent)
        self.assertIsInstance(ext_str, str)
        self.assertTrue("files" in ext_str)
        ext_str = build_extension_count_str(self.invalid_path)
        self.assertIsInstance(ext_str, str)
        self.assertEqual(ext_str, 'default')
        self.assertFalse("files" in ext_str)

    def test_bytes_to_readable(self):
        self.assertEqual(bytes_to_readable('string'), '0 bytes')
        self.assertTrue("ERROR:" in str(bytes_to_readable(-999)))
        self.assertEqual(bytes_to_readable(0), '0 bytes')
        if self.is_win:
            byte = 1024
        else:
            byte = 1000
        self.assertEqual(bytes_to_readable(byte ** 1), '01.00 KiB')
        self.assertEqual(bytes_to_readable(byte ** 2), '01.00 MiB')
        self.assertEqual(bytes_to_readable(byte ** 3), '01.00 GiB')
        self.assertEqual(bytes_to_readable(byte ** 4), '01.00 TiB')

    def test_split_path(self):
        parent, curr_dir, file_name, file_ext = split_path(self.valid_dir)
        self.assertEqual(str(self.valid_dir.parent), parent)
        self.assertEqual(self.valid_dir.parts[-1], curr_dir)
        self.assertEqual(self.valid_dir.name, file_name)
        self.assertTrue(self.valid_dir.is_dir())
        self.assertFalse(self.valid_dir.is_file())
        self.assertTrue('.' not in file_ext)
        parent, curr_dir, file_name, file_ext = split_path(self.valid_file)
        self.assertEqual(str(self.valid_file.parent), parent)
        self.assertEqual(self.valid_file.parts[-1], curr_dir)
        self.assertEqual(self.valid_file.name, file_name + file_ext)
        self.assertTrue(self.valid_file.is_file())
        self.assertFalse(self.valid_file.is_dir())
        self.assertTrue('.' in file_ext)
        parent, curr_dir, file_name, file_ext = split_path(self.noext_file)
        self.assertEqual(str(self.noext_file.parent), parent)
        self.assertEqual(self.noext_file.parts[-1], curr_dir)
        self.assertEqual(self.noext_file.name, file_name)
        self.assertTrue(self.noext_file.is_file())
        self.assertFalse(self.noext_file.is_dir())
        self.assertTrue('.' not in file_ext)

    def test_save_output_txt(self):
        status = save_output_txt(self.out_path,
                                 'out_1.log', 'test output1',
                                 delim_tag=True, replace_ext=False)
        self.assertTrue("SUCCESS" in status)
        self.assertTrue(
            os.path.exists(os.path.join(self.out_path, '~out_1.log')))
        status = save_output_txt(self.out_path,
                                 'out_2.txt', 'test output2',
                                 delim_tag=False, replace_ext=True)
        self.assertTrue("SUCCESS" in status)
        self.assertTrue(
            os.path.exists(os.path.join(self.out_path, 'out_2.txt')))
        status = save_output_txt(self.out_path,
                                 'out_3.log', 'test output3',
                                 delim_tag=False, replace_ext=False)
        self.assertTrue("SUCCESS" in status)
        self.assertTrue(
            os.path.exists(os.path.join(self.out_path, 'out_3.log')))

    def test_get_sha256_hash(self):
        sha_hex = get_sha256_hash(self.valid_file)
        self.assertIsInstance(sha_hex, str)
        self.assertEqual(len(sha_hex), 64)

    def test_get_dir_stats(self):
        dir_stats = get_dir_stats(self.valid_dir)
        self.assertIsInstance(dir_stats, list)
        self.assertEqual(len(dir_stats), len(os.listdir(self.valid_dir)))
        for d_stat in dir_stats:
            self.assertIsInstance(d_stat, list)
            self.assertIsInstance(d_stat[0], str)

    def test_get_files(self):
        file_list = get_files(self.valid_dir, file_ext='.dcm')
        for _file in file_list:
            self.assertTrue(_file.is_file())
            self.assertIsInstance(_file, pathlib.Path)
            self.assertIn('.dcm', str(_file))
            self.assertNotIn('.py', str(_file))

    def test_get_directories(self):
        dir_list = get_directories(self.valid_dir)
        for folder in dir_list:
            self.assertTrue(folder.is_dir())
            self.assertIsInstance(folder, pathlib.Path)
            self.assertEqual(len(dir_list),
                             len(os.listdir(self.valid_dir)))

    def test_get_extensions(self):
        ext_list = get_extensions(self.valid_dir)
        if ext_list:
            for file_ext in ext_list:
                self.assertIsInstance(file_ext, str)
                self.assertIn(file_ext, self.valid_ext)
            self.assertEqual(len(ext_list), 1)

    def tearDown(self) -> None:
        if os.path.exists(self.out_path):
            shutil.rmtree(self.out_path)

    if __name__ == '__main__':
        unittest.main()
