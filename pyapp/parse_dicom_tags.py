# -*- coding: UTF-8 -*-
"""Module to read and parse DICOM tag data from text files."""
import argparse
import inspect
import math
import os
import string
import sys
import time
from collections import OrderedDict
import xlsxwriter
from pylibs import config
from pylibs import file_tools
from pylibs import dicom_tools


BASE_DIR, SCRIPT_NAME = os.path.split(os.path.abspath(__file__))
PARENT_PATH, CURR_DIR = os.path.split(BASE_DIR)
IS_WINDOWS = sys.platform.startswith('win')

MAX_EXCEL_TAB_NODATE = 22  # '_12212019' = 9 chars
MAX_EXCEL_TAB_DIR = 27
MAX_EXCEL_TAB = 31

ALPHABET = string.ascii_uppercase
valid_chars = f"-_.()~{ALPHABET}{string.digits}"


def get_header_column_widths(input_tag_list: list) -> dict:
    """Returns dynamically sized column widths based on cell values."""
    # list: [row1:[hdr1, ..., hdrN], row2:[data1, ..., dataN]... rowN]
    headers = list(input_tag_list[0])
    scalar = 1.2  # account for presentations difference
    hdr_col_width_dict = OrderedDict([(hdr, -1) for hdr in headers])
    for row_num, tag_list in enumerate(input_tag_list):
        for col_num, cell_val in enumerate(tag_list):
            header = headers[col_num]
            max_length = len(cell_val)
            if hdr_col_width_dict[header] < max_length:
                if max_length > 10:
                    max_length *= scalar
                # each char, +2 for readability
                hdr_col_width_dict[header] = int(math.ceil(max_length)) + 2
    if config.VERBOSE:
        print("\ndynamically sized columns widths:")
        for key, value in hdr_col_width_dict.items():
            print(f"   {key:28} \t {value} chars")
    return hdr_col_width_dict


def export_to_excel(output_path: str, filename: str, stat_list: list) -> str:
    """Exports DICOM tag data into output Excel report file with markup."""
    def_name = inspect.currentframe().f_code.co_name
    status_str = f"{def_name}() in: '{output_path}'\n"
    print(status_str)
    if len(stat_list) > 0:
        try:
            file_basename, file_ext = filename.split('.')
            output_filepath = os.path.join(output_path, filename)
            workbook = xlsxwriter.Workbook(f"{output_filepath}")
            ws1 = workbook.add_worksheet(
                file_basename[:MAX_EXCEL_TAB])  # <= 31 chars
            ws1.freeze_panes(1, 0)
            # Add formatting: RED fill with dark red text
            format_red = workbook.add_format({'bg_color': '#FFC7CE',
                                              'font_color': '#9C0006',
                                              'bold': False})
            # Add formatting: GREEN fill with dark green text
            format_green = workbook.add_format({'bg_color': '#C6EFCE',
                                                'font_color': '#006100',
                                                'bold': False})
            # includes both header and cell values in calculation
            hdr_col_width_dict = get_header_column_widths(stat_list)
            xls_font_name = 'Segoe UI'
            font_pt_size = 11  # default: 11 pt
            header_format = workbook.add_format({'bold': True,
                                                 'underline': True,
                                                 'font_color': 'blue',
                                                 'center_across': True})
            header_format.set_font_size(font_pt_size)
            header_format.set_font_name(xls_font_name)
            stat_list.pop(0)  # remove header row
            for idx, key_hdr in enumerate(hdr_col_width_dict):
                alpha = ALPHABET[idx]
                col_width_val = hdr_col_width_dict[key_hdr]
                ws1.set_column(f"{alpha}:{alpha}", col_width_val)  # all rows
                ws1.write(f"{alpha}1", f"{key_hdr}:", header_format)
            ctr_int = workbook.add_format()
            ctr_int.set_num_format('0')
            ctr_int.set_align('vcenter')
            ctr_int.set_align('center')
            ctr_txt = workbook.add_format()
            ctr_txt.set_align('vcenter')
            ctr_txt.set_align('center')
            ctr_txt.set_font_name(xls_font_name)
            ctr_date = workbook.add_format()
            ctr_date.set_num_format('mm/dd/yy hh:mm AM/PM')
            ctr_date.set_align('vcenter')
            ctr_date.set_align('center')
            ctr_date.set_font_name(xls_font_name)
            if stat_list:  # if data after popped header row
                last_alpha = ALPHABET[
                    len(stat_list[0]) - 1]  # last index -1 of len()
                ws1.autofilter(f"A1:{last_alpha}65536")
                # A0, B1, C2, D3, E4, F5, G6, H7 = 8 indexed entries
                num = 1  # first row in Excel
                if len(stat_list) > 0:
                    for idx, tag_list in enumerate(stat_list):
                        num += 1  # not header row
                        for count, tag_val in enumerate(tag_list):
                            alpha = ALPHABET[count]
                            ws1.write(f"{alpha}{num}", str(tag_val),
                                      ctr_txt)
            ws1.conditional_format('F2:F%d' % num, {'type': 'text',
                                                    'criteria': 'containing',
                                                    'value': 'Physicists',
                                                    'format': format_green})
            ws1.conditional_format('C2:C%d' % num, {'type': 'text',
                                                    'criteria': 'containing',
                                                    'value': 'OT',
                                                    'format': format_red})
            workbook.close()
            status_str = f"SUCCESS! {def_name}() \n{output_filepath}"
        except Exception as exp:
            status_str = f"~!ERROR!~ {def_name}() {sys.exc_info()[0]}\n{exp}"
        finally:
            return status_str


def is_fuji_tag_dump(txt_file_lines: list) -> tuple:
    """Dynamically determines if input is Fuji or DCMTK."""
    isFuji = False
    isDCMTK = False
    # check only first n-lines of input file for unique substring match
    for line_str in txt_file_lines[0:5]:
        if dicom_tools.FUJI_TAG in line_str:
            isFuji = True
        if dicom_tools.DCMTK_TAG in line_str:
            isDCMTK = True
    return isFuji, isDCMTK


def get_tag_line_number(tag_keyword: str = '(0008,0020)',
                        lines: list = []) -> int:
    """Optimization: stop iterating once index to tag_keyword is located."""
    # using tag_keyword '(0008,0050)' or '0008 0050'
    # if tag_keyword not in input_lines, return -1.
    return next((line_num for line_num, line_str in enumerate(lines)
                 if tag_keyword in line_str), -1)


def get_tag_indices(tags: list, lines: list,
                    is_optimized: bool = True) -> dict:
    """Iterates through lines to find desired DICOM tags."""
    tag_indices_dict = OrderedDict([(hdr, '') for hdr in tags])
    # using tag_keyword '(0008,0050)' or '0008 0050'
    if is_optimized:
        for tag_keyword in tags:
            line_num = get_tag_line_number(tag_keyword, lines)
            if line_num != -1:
                tag_indices_dict[tag_keyword] = lines[line_num]
    else:
        for line_num, line_str in enumerate(lines):
            for tag_keyword in tags:
                if tag_keyword in line_str:
                    tag_indices_dict[tag_keyword] = line_str
    if config.DEBUG:
        for tag_key, tag_value in tag_indices_dict.items():
            print(f"{tag_key}={tag_value}", end='')
    return tag_indices_dict


def parse_dicom_tag_dump(input_headers: list, input_path: str) -> list:
    """Parse DICOM desired tag data from input .txt files."""
    def_name = inspect.currentframe().f_code.co_name
    status_str = f"{def_name}() in: '{input_path}'\n"
    print(status_str)
    tag_name = 'tagdump'
    source_ext = '.txt'
    filename_list, filepath_list = file_tools.get_files(input_path, source_ext)
    file_count = 0
    dump_count = 0
    output_tag_list = [input_headers]  # first row contains headers
    if not filepath_list:
        error_msg = f"~!ERROR!~ missing files, check path: \n{input_path}"
        print(error_msg)
    else:
        print(f"PARSING: ({len(filepath_list)}) '{source_ext}' files")
        for this_file in filepath_list:
            file_count += 1
            parsed_file_list = []
            read_file_handle = open(this_file, 'r')
            if tag_name not in this_file:
                print(f"   reading_{file_count:03}: {this_file}")
                lines_list = read_file_handle.readlines()
                # dynamically determine which input file format:
                (isFuji, isDCMTK) = is_fuji_tag_dump(lines_list[0:5])
                if isFuji:
                    elements = dicom_tools.build_fuji_tag_dict(this_file)
                elif isDCMTK:
                    elements = dicom_tools.build_dcmtk_tag_dict(this_file)
                else:
                    elements = None  # input '.txt' not a tag dump
                if elements:
                    # re-initializes output dict for each file to blank values
                    tag_dict = OrderedDict([(hdr, '')
                                            for hdr in list(elements.keys())])
                    path_head, path_tail = os.path.split(this_file)
                    tag_dict['filename'] = path_tail
                    # using values: tag '(0008,0050)' or '0008 0050'
                    tag_indices = get_tag_indices(list(elements.values()),
                                                  lines_list)
                    tag_num = 0
                    dump_count += 1
                    for tag_key, tag_value in elements.items():
                        tag_num += 1
                        line_str = tag_indices[tag_value]
                        if len(tag_indices) > 0:
                            if isDCMTK:
                                # parse value between square brackets [..]
                                if '[' in line_str:
                                    target_value = \
                                        line_str.split('[', 1)[1].split(']')[0]
                                    tag_dict[tag_key] = target_value
                                elif '=' in line_str:
                                    target_value = \
                                        line_str.split('=', 1)[1].split('#')[0]
                                    tag_dict[tag_key] = target_value.strip()
                            elif isFuji:
                                # parse value between double quotes "..."
                                if '"' in line_str:
                                    target_value = \
                                        line_str.split('"', 1)[1].split('"')[0]
                                    tag_dict[tag_key] = target_value
                        if config.DEBUG:
                            print(f"tag_{tag_num:02} {tag_key:24} "
                                  f"\t{tag_value} line: {line_str:40} "
                                  f"len:{len(line_str):02} chars")
                    for tag_attribute, parsed_val in tag_dict.items():
                        parsed_file_list.append(parsed_val)
                    output_tag_list.append(parsed_file_list)
            read_file_handle.close()
        print(
            f"EXTRACTION: {dump_count} dumps of "
            f"{file_count} '{source_ext}' files")
    return output_tag_list


def get_cmd_args() -> str:
    """Command line input on directory to scan recursively for DICOM dumps."""
    def_name = inspect.currentframe().f_code.co_name
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input", type=str, help="input path")
    args = parser.parse_args()
    if args.input is None:
        if config.DEMO_ENABLED:
            input_path = os.path.join(PARENT_PATH, 'input', 'tag_dumps')
        else:
            input_path = os.path.join(PARENT_PATH, 'tag_dumps_all')
    else:
        input_path = args.input
        if os.path.exists(input_path) and os.path.isdir(input_path):
            print(f"{def_name}() dumping path:'{input_path}'")
        else:
            parser.error(f"invalid path: '{input_path}'")
            input_path = None
    return input_path


def main():
    """Driver to read and parse DICOM tag data from text files."""
    print(f"{SCRIPT_NAME} starting...")
    start = time.perf_counter()
    config.print_header(SCRIPT_NAME)
    input_path = get_cmd_args()
    if os.path.exists(input_path):
        if config.DEMO_ENABLED:
            output_path = os.path.join(PARENT_PATH, 'output')
        else:
            output_path = os.path.join(PARENT_PATH, CURR_DIR, 'tag_dumps_all')
        if not os.path.exists(output_path):
            os.makedirs(output_path)
        all_tag_list = parse_dicom_tag_dump(dicom_tools.headers, input_path)
        curr_date, curr_time = file_tools.generate_date_str()
        filename = f"{config.TEMP_TAG}dicom_tag_dumps.xlsx"
        # works on both linux and windows
        if len(all_tag_list) > 1:  # more than just headers
            xls_status = export_to_excel(output_path, filename, all_tag_list)
            print(xls_status)
    else:
        print(f"~!ERROR!~ invalid path: {input_path}")
    end = time.perf_counter() - start
    print(f"{SCRIPT_NAME} finished in {end:0.2f} seconds")


if __name__ == "__main__":
    main()
