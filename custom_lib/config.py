#!python3
# -*- coding: utf-8 -*-
import json
import locale
import os
import platform as p
import sys
from urllib.error import HTTPError
from urllib.error import URLError
from urllib.request import Request
from urllib.request import urlopen

BASE_DIR, SCRIPT_NAME = os.path.split(os.path.abspath(__file__))
PARENT_PATH, CURR_DIR = os.path.split(BASE_DIR)
IS_WINDOWS = sys.platform.startswith('win')
DEBUG = False
VERBOSE = False
DEMO_ENABLED = True
SHOW_METHODS = False
TEMP_TAG = '~'


def show_method(method_name: str) -> None:
    if SHOW_METHODS:
        print(f"{method_name.upper()}()")


def get_isp_info() -> str:
    """Get current ISP information of connected host"""
    isp_req = Request('http://ipinfo.io/json')
    try:
        response = urlopen(isp_req)
    except HTTPError as exception:
        print(f"HTTPError code: {exception.code}")
    except URLError as exception:
        print(f"\nERROR: {sys.exc_info()[0]} {exception}")
        return exception.reason
    isp_data = json.load(response)
    info_str = json.dumps(isp_data, sort_keys=True, indent=4)
    return info_str


__author__ = "averille"
__email__ = "github.pdx@runbox.com"
__status__ = "demo"
__license__ = "MIT"
__version__ = "1.3.0"
__login_user__ = os.getlogin()
__python_version__ = f"{p.python_version()}"
__host_arch__ = f"{p.system()} {p.architecture()[0]} {p.machine()}"
__sys_encoding__ = str(sys.getfilesystemencoding()).upper()
__script_encoding__ = f"{SCRIPT_NAME:26} {__sys_encoding__}"
__host_enc__ = locale.getpreferredencoding()
__host_info__ = f"{p.node():12} {__host_enc__} {__host_arch__:16}"
__header__ = f"""  license: \t{__license__}
  python:  \t{__python_version__}
  host:    \t{__host_info__}
  login:   \t{__login_user__}
  script:  \t{__script_encoding__}
  author:  \t{__author__}
  email:   \t{__email__}
  status:  \t{__status__}
  version: \t{__version__}
  isp_info:\t{get_isp_info()}
"""


def print_header() -> None:
    """Display project status, host characteristics, and ISP information"""
    print(__header__)
