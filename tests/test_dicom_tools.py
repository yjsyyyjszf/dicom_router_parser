import unittest
import os

BASE_DIR, SCRIPT_NAME = os.path.split(os.path.abspath(__file__))
PARENT_PATH, CURR_DIR = os.path.split(BASE_DIR)


class TestDicomTools(unittest.TestCase):
    def test_extract_tag(self):
        self.assertNotEqual('this', 'that')
    # TO DO: 


if __name__ == '__main__':
    unittest.main()
