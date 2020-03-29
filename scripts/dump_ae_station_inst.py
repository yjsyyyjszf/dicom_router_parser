#!python3
# -*- coding: utf-8 -*-
from collections import OrderedDict
import datetime
import json
import locale
import math
import optparse
import os
import platform
import string
import sys
import time
from urllib.error import HTTPError
from urllib.error import URLError
from urllib.request import Request
from urllib.request import urlopen
import xlsxwriter


BASE_DIR, SCRIPT_NAME = os.path.split(os.path.abspath(__file__))
PARENT_PATH, CURR_DIR = os.path.split(BASE_DIR)

DEBUG = False
VERBOSE = False
DEMO = True


def get_isp_info() -> str:
    isp_req = Request('http://ipinfo.io/json')
    try:
        response = urlopen(isp_req)
    except HTTPError as exception:
        print('HTTPError code: ', exception.code)
    except URLError as exception:
        print(f"\nERROR: {sys.exc_info()[0]} {exception}")
        return exception.reason
    isp_data = json.load(response)
    info_str = json.dumps(isp_data, sort_keys=True, indent=4)
    return info_str


__author__ = "Emile Averill"
__email__ = "github.pdx@runbox.com"
__status__ = "demo"
__license__ = "MIT"
__version__ = "1.2.6"
__python_version__ = f"{platform.python_version()}"
__host_architecture__ = f"{platform.system()} {platform.architecture()[0]} {platform.machine()}"
__script_encoding__ = f"{SCRIPT_NAME:26} ({str(sys.getfilesystemencoding()).upper()})"
__host_info__ = f"{platform.node():26} ({locale.getpreferredencoding()}) {__host_architecture__:16}"
__header__ = f"""  license: \t{__license__}
  python:  \t{__python_version__}
  host:    \t{__host_info__}
  script:  \t{__script_encoding__}
  author:  \t{__author__}
  email:   \t{__email__}
  status:  \t{__status__}
  version: \t{__version__}
  isp_info:\t{get_isp_info()}
"""


def print_header() -> None:
    print(__header__)


MAX_EXCEL_TAB_NODATE = 22  # '_12212019' = 9 chars
MAX_EXCEL_TAB_DIR = 27
MAX_EXCEL_TAB = 31
TEMP_TAG = '~'
ALPHABET = string.ascii_uppercase
valid_chars = f"-_.()~{string.ascii_letters}{string.digits}"

AVOID_DIRS = []
AVOID_DIRS.append('.git')
AVOID_DIRS.append('.idea')
AVOID_DIRS.append('venv')
AVOID_DIRS.append('__pycache__')


def sanitize_input(input_str: str = 'default'):
    sanitized = input_str
    invalid_chars = "{}[]()<>^#+*?$@&%$!,:;/\\ "  # keep - and .
    sanitized = sanitized.replace("-", "")
    for char in invalid_chars:
        if char in input_str:
            sanitized = sanitized.replace(char, "")
    return sanitized


def generate_date_str() -> tuple:
    now = datetime.datetime.now()
    date = now.strftime("%m-%d-%Y")
    time = now.strftime("%H%M%p").lower()
    return (date, time)


def get_header_column_widths(input_tag_list: list = []) -> dict:
    # list of lists: [row1:[hdr1, hdr2, ..., hdrN], row2:[data1, data2, ..., dataN]... rowN]
    headers = list(input_tag_list[0])
    scalar = 1.2  # account for presentations difference in Excel based on font
    hdr_col_width_dict = OrderedDict([(hdr, -1) for hdr in headers])  # init to -1
    for row_num, tag_list in enumerate(input_tag_list):
        for col_num, cell_val in enumerate(tag_list):
            header = headers[col_num]
            max_length = len(cell_val)
            if (hdr_col_width_dict[header] < max_length):
                if max_length > 10:
                    max_length *= scalar
                hdr_col_width_dict[header] = int(math.ceil(max_length)) + 2  # each char=1, +2 for readability
    if (VERBOSE):
        print("\nDynamically sized columns widths:")
        for (key, value) in hdr_col_width_dict.items():
            print(f"   {key:28} \t {value} chars")
    return hdr_col_width_dict


def export_to_excel(output_path: str = 'default_path', output_filename: str = 'default_filename', stat_list_of_lists: list = []) -> str:
    def_name = sys._getframe().f_code.co_name.upper()
    status_str = f"{def_name}() in: '{output_path}'\n"
    print(status_str)
    if (len(stat_list_of_lists) > 0):
        try:
            file_basename, file_ext = output_filename.split('.')
            output_filepath = os.path.join(output_path, output_filename)
            workbook = xlsxwriter.Workbook(f"{output_filepath}")
            worksheet1 = workbook.add_worksheet(file_basename[:MAX_EXCEL_TAB])  # <= 31 chars
            worksheet1.freeze_panes(1, 0)
            
            # https://xlsxwriter.readthedocs.io/example_conditional_format.html#ex-cond-format
            # add formating: light red fill with dark red text
            # format_red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': False})
            # add formating: green fill with dark green text
            # format_green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': False})
            
            hdr_col_width_dict = get_header_column_widths(stat_list_of_lists)   # includes both header and cell values in calcuation

            xls_font_name = 'Segoe UI'
            font_pt_size = 11   # default is 11 pt
            header_format = workbook.add_format({'bold': True, 'underline': True, 'font_color': 'blue', 'center_across': True})
            header_format.set_font_size(font_pt_size)
            header_format.set_font_name(xls_font_name)
            
            stat_list_of_lists.pop(0)  # remove header row
            for idx, key_hdr in enumerate(hdr_col_width_dict):
                alpha = ALPHABET[idx]
                col_width_val = hdr_col_width_dict[key_hdr]
                worksheet1.set_column(f"{alpha}:{alpha}", col_width_val)  # all rows
                worksheet1.write(f"{alpha}1", f"{key_hdr}:", header_format)
            
            centered_cells_int = workbook.add_format()
            centered_cells_int.set_num_format('0')
            centered_cells_int.set_align('vcenter')
            centered_cells_int.set_align('center')
            centered_cells = workbook.add_format()
            centered_cells.set_align('vcenter')
            centered_cells.set_align('center')
            centered_cells.set_font_name(xls_font_name)
             
            date_cell_format = workbook.add_format()
            date_cell_format.set_num_format('mm/dd/yy hh:mm AM/PM')
            date_cell_format.set_align('vcenter')
            date_cell_format.set_align('center')
            date_cell_format.set_font_name(xls_font_name)

            if (stat_list_of_lists):  # if data after popped header row
                last_alpha = ALPHABET[len(stat_list_of_lists[0]) - 1]  # last index -1 of len()
                worksheet1.autofilter(f"A1:{last_alpha}65536")
                # A0, B1, C2, D3, E4, F5, G6, H7 = 8 indexed entries
                row_num = 1  # first row in Excel
                if (len(stat_list_of_lists) > 0):
                    for idx, tag_list in enumerate(stat_list_of_lists):
                        row_num += 1  # not header row
                        for count, tag_val in enumerate(tag_list):
                            alpha = ALPHABET[count]
                            worksheet1.write(f"{alpha}{row_num}", str(tag_val), centered_cells)
            workbook.close()
            status_str = f"SUCCESS! {def_name}() \n{output_filepath}"
        except (Exception) as exception:
            status_str = f"~!ERROR!~ {def_name}() {sys.exc_info()[0]}\n{exception}"
        finally:
            return status_str


transfer_syntaxes = OrderedDict([("1.2.840.10008.1.2", 'ImplicitVRLittleEndian'),    # ILE ImplicitVRLittleEndian
                               ("1.2.840.10008.1.2.1", 'ExplicitVRLittleEndian'),    # ELE ExplicitVRLittleEndian
                               ("1.2.840.10008.1.2.2", 'ExplicitVRBigEndian'),       # EBE ExplicitVRBigEndian
                               ("1.2.840.10008.1.2.4.50", 'JPEGBaselineProcess1'),   # JPG1 JPEGBaselineProcess1
                               ("1.2.840.10008.1.2.4.51", 'JPEGBaselineProcess2'),   # JPG2  JPEGBaselineProcess2
                               ("1.2.840.10008.1.2.4.57", 'JPEGLossless14'),         # JPG14  JPEGLossless14
                               ("1.2.840.10008.1.2.4.70", 'JPEGLossless14FOP'),      # JPG14  JPEGLossless14FOP
                               ("1.2.840.10008.1.2.4.90", 'JPEG2000Lossless'),       # J2KL  JPEG2000LosslessOnly
                               ("1.2.840.10008.1.2.4.91", 'JPEG2000'),               # J2K  JPEG2000
                               ("1.2.840.10008.1.2.5", 'RunLengthEncoding')])        # RLE  RunLengthEncodingLossless


'''
   0002 0016 | sourceApplicationEntityTitle
   0008 0050 | accessionNumber
   0008 0060 | modality
   0008 1010 | stationName
   0008 0080 | institutionName
   0008 0070 | manufacturer
   0008 1090 | manufacturerModelName

http://dicomlookup.com/default.asp
http://dicom.nema.org/dicom/2013/output/chtml/part05/sect_6.2.html
'CS' refers to the value representation: Code String 16 bytes max

FUJI:  vertical bars | or ---- present in each line
key: tag with no parentheses or comma '#### ####'
value: between double quotes "..."
example:   0008 0060 | modality       | CS |     1 | "CR"

DCMTK:  number sign # present in each line
key: between (####,####)
value: between square brackets [...]
example:   (0008,0060) CS [CT]         #   2, 1 Modality
'''


# tag: (0008,0050) is represented as '0008 0050' for FUJI sourced files
def build_fuji_tag_dict(input_filename: str = 'default_filename') -> dict:
    fuji_tag_dict = OrderedDict()
    fuji_tag_dict['filename'] = input_filename
    fuji_tag_dict['accessionNumber'] = '0008 0050'
    fuji_tag_dict['modality'] = '0008 0060'
    fuji_tag_dict['sourceApplicationEntityTitle'] = '0002 0016'
    fuji_tag_dict['stationName'] = '0008 1010'
    fuji_tag_dict['institutionName'] = '0008 0080'
    fuji_tag_dict['manufacturer'] = '0008 0070'
    fuji_tag_dict['manufacturerModelName'] = '0008 1090'
    fuji_tag_dict['transferSyntaxUid'] = '0002 0010'
    return fuji_tag_dict


# tag: (0008,0050) is represented as '(0008,0050)' for DCMTK sourced files
def build_dcmtk_tag_dict(input_filename: str = 'default_filename') -> dict:
    dcmtk_tag_dict = OrderedDict()
    dcmtk_tag_dict['filename'] = input_filename
    dcmtk_tag_dict['accessionNumber'] = '(0008,0050)'
    dcmtk_tag_dict['modality'] = '(0008,0060)'
    dcmtk_tag_dict['sourceApplicationEntityTitle'] = '(0002,0016)'
    dcmtk_tag_dict['stationName'] = '(0008,1010)'
    dcmtk_tag_dict['institutionName'] = '(0008,0080)'
    dcmtk_tag_dict['manufacturer'] = '(0008,0070)'
    dcmtk_tag_dict['manufacturerModelName'] = '(0008,1090)'
    dcmtk_tag_dict['transferSyntaxUid'] = '(0002,0010)'
    return dcmtk_tag_dict


def is_fuji_tag_dump(txt_file_lines: list = []) -> tuple:
    isFuji = False
    isDCMTK = False
    # check only first n-lines of input file for unique substring match
    for line_str in txt_file_lines[0:5]:
        if ('Grp  Elmt | Description' in line_str):
            isFuji = True
        if ('Dicom-Meta-Information-Header' in line_str):
            isDCMTK = True
        # print(f"line: '{line_str[0:48]}'\n   isFuji: {isFuji}  isDCMTK: {isDCMTK}")
    return (isFuji, isDCMTK)


def get_tag_line_number(tag_keyword: str = '(0008,0020)', input_lines: list = []) -> int:
    # optimization: stop iterating once index to tag_keyword is located
    # using tag_keyword '(0008,0050)' or '0008 0050'
    return next((line_num for line_num, string in enumerate(input_lines) if tag_keyword in string), -1)
    # if tag_keyword not in input_lines, return -1.


def get_tag_indices(tag_keywords: list = [], input_lines: list = [], isOptimized: bool = True) -> dict:
    tag_indices_dict = OrderedDict([(hdr, '') for hdr in tag_keywords])
    # using tag_keyword '(0008,0050)' or '0008 0050'
    if (isOptimized):
        for tag_keyword in tag_keywords:
            line_num = get_tag_line_number(tag_keyword, input_lines)
            if (line_num != -1):
                tag_indices_dict[tag_keyword] = input_lines[line_num]
    else:
        # interate through all lines in list to find desired tags
        for line_num, line_str in enumerate(input_lines):
            for tag_keyword in tag_keywords:
                if (tag_keyword in line_str):
                    tag_indices_dict[tag_keyword] = line_str
    if (DEBUG):
        for tag_key, tag_value in tag_indices_dict.items():
            print(f"{tag_key}={tag_value}")
    return tag_indices_dict


def check_pathname_skip(input_path: str = 'default_path') -> bool:
    if (input_path):
        path_hit = next((s for s in AVOID_DIRS if s in input_path), True)
        if (path_hit):
            return True
        else:
            return False


def find_files_wext(input_path: str = 'default_path', input_ext: str = '.txt') -> tuple:
    def_name = sys._getframe().f_code.co_name.upper()
    print(f"{def_name}({input_ext})")
    file_list = []
    file_path_list = []
    # get files, paths for SPECIFIC file extention (including sub-folders)
    if (os.path.exists(input_path)):
        for root_path, dirs, files in os.walk(input_path):
            for this_file in files:
                if this_file.lower().endswith((input_ext)):
                    this_file_path = os.path.join(root_path, this_file)
                    if check_pathname_skip(this_file_path):
                        file_list.append(this_file)
                        file_path_list.append(this_file_path)
        file_path_list.sort(key=lambda x: (-x.count(os.sep), x))  # sort based on filename
    return sorted(file_list), file_path_list


def parse_dicom_tag_dump(input_headers: list = [], input_folder_path: str = 'default_path') -> list:
    def_name = sys._getframe().f_code.co_name.upper()
    status_str = f"{def_name}() in: '{input_folder_path}'\n"
    print(status_str)
    tag_name = 'tagdump'
    source_ext = '.txt'
    filename_list, filepath_list = find_files_wext(input_folder_path, source_ext)
    file_count = 0
    dump_count = 0
    output_tag_list = [input_headers]  # first row contains headers
   
    if not filepath_list:
        error_msg = f"~!ERROR!~ missing input files, check input path: \n{input_folder_path}"
        print(error_msg)
    else:
        print(f"PARSING: ({len(filepath_list)}) '{source_ext}' files")
        for this_file in filepath_list:
            file_count += 1
            parsed_file_list = []
            read_file_handle = open(this_file, 'r')
            
            if (tag_name not in this_file):
                print(f"   reading_{file_count:03}: {this_file}")
                lines_list = read_file_handle.readlines()
                
                # dynamically determine which input file format: based on unique substring in first n-lines
                (isFuji, isDCMTK) = is_fuji_tag_dump(lines_list[0:5])
                if (isFuji):
                    target_element_dict = build_fuji_tag_dict(this_file)
                elif (isDCMTK):
                    target_element_dict = build_dcmtk_tag_dict(this_file)
                else:
                    target_element_dict = None  # input '.txt' is not a tag dump

                if (target_element_dict):
                    # re-initializes output dict for each file to blank values
                    tag_dict = OrderedDict([(hdr, '') for hdr in list(target_element_dict.keys())])
                    path_head, path_tail = os.path.split(this_file)
                    tag_dict['filename'] = path_tail
                    
                    # using values: tag '(0008,0050)' or '0008 0050'
                    tag_indices = get_tag_indices(list(target_element_dict.values()), lines_list)
                    tag_num = 0
                    dump_count += 1
                    for tag_key, tag_value in target_element_dict.items():
                        tag_num += 1
                        line_str = tag_indices[tag_value]
                        if (len(tag_indices) > 0):
                            if (isDCMTK):
                                if '[' in line_str:  # get value between square brackets [..]
                                    target_value = line_str.split('[', 1)[1].split(']')[0]
                                    tag_dict[tag_key] = target_value
                            elif (isFuji):
                                if '"' in line_str:  # get value between double quotes "..."
                                    target_value = line_str.split('"', 1)[1].split('"')[0]
                                    tag_dict[tag_key] = target_value
                        if (DEBUG):
                            print(f"tag_{tag_num:02} {tag_key:24} \t{tag_value}  line: {line_str:40} len:{len(line_str):02} chars")
                        
                    for tag_attribute, parsed_val in tag_dict.items():
                        parsed_file_list.append(parsed_val)
                    output_tag_list.append(parsed_file_list)
                    
            read_file_handle.close()
        print(f"EXTRACTION: {dump_count} dumps of {file_count} '{source_ext}' files")
    return output_tag_list


def get_cmd_args():
    def_name = sys._getframe().f_code.co_name.upper()
    parser = optparse.OptionParser()
    parser.add_option("-i", "--input", help="input path")
    options, remain = parser.parse_args()
    if options.input is None:
        if (DEMO):
            input_path = os.path.join(PARENT_PATH, 'input', 'tag_dumps')
        else:
            input_path = os.path.join(PARENT_PATH, CURR_DIR, 'tag_dumps_all')
    else:
        input_path = options.input
        if (os.path.exists(input_path) and os.path.isdir(input_path)):
            print(f"{def_name}() dumping path:'{input_path}'")
        else:
            parser.error(f"invalid path: '{input_path}'")
            input_path = None
    return input_path
    
    
if __name__ == "__main__":
    print(f"{SCRIPT_NAME} starting...")
    start = time.perf_counter()
    print_header()

    input_path = get_cmd_args()
    if os.path.exists(input_path):
        if (DEMO):
            output_path = os.path.join(PARENT_PATH, 'output')
        else:
            output_path = os.path.join(PARENT_PATH, CURR_DIR, 'tag_dumps_all')
        if not os.path.exists(output_path):
            os.makedirs(output_path)
            
        headers = ["filename", "accessionNumber", "modality", "sourceApplicationEntityTitle", "stationName",
                   "institutionName", "manufacturer", "manufacturerModelName", "transferSyntaxUid"]
        all_tag_list = parse_dicom_tag_dump(headers, input_path)

        curr_date, curr_time = generate_date_str()
        output_filename = f"{TEMP_TAG}dicom_tag_dumps.xlsx"
        
        # works on both Linux and Windows
        if (len(all_tag_list) > 1):  # more than just headers
            xls_status = export_to_excel(output_path, output_filename, all_tag_list)
            print(xls_status)
    else:
        print(f"~!ERROR!~ invalid path: {input_path}")
        
    end = time.perf_counter() - start
    print(f"{SCRIPT_NAME} finished in {end:0.2f} seconds")
