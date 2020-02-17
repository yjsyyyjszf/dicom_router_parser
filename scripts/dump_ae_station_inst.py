#!python3
import glob
import sys
import os
import platform
import locale
import time
import datetime
import re
import zlib
import hashlib
import string
import xlsxwriter
from collections import OrderedDict
import json
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

BASE_DIR, SCRIPT_NAME = os.path.split(os.path.abspath(__file__))
DEBUG = False
VERBOSE = False

def get_isp_info() -> str:
    isp_req = Request('http://ipinfo.io/json')
    try:
        response = urlopen(isp_req)
    except HTTPError as exception:
        print('HTTPError code: ', exception.code)
    except URLError as exception:
        print("\nERROR: {} {}".format(sys.exc_info()[0], exception))
        return exception.reason
    isp_data = json.load(response)
    info_str = json.dumps(isp_data, sort_keys=True, indent=4)
    return info_str

__author__ = "Emile Averill"
__email__  = "github.pdx@runbox.com"
__status__ = "testing"
__license__ = "MIT"
__version__ = "1.2.5"
__python_version__ = "{}".format(platform.python_version())
__host_architecture__ = "{} {} {}".format( platform.system(), platform.architecture()[0], platform.machine())
__script_encoding__ = "{:26} ({})".format(SCRIPT_NAME, str(sys.getfilesystemencoding()).upper())
__host_info__ = "{:26} ({}) {:16}".format(platform.node(), locale.getpreferredencoding(), __host_architecture__)
__header__ = """  license: \t{}
  python:  \t{}
  host:    \t{}
  script:  \t{}
  author:  \t{} 
  email:   \t{}
  status:  \t{}
  version: \t{}
{}
""".format(__license__, __python_version__, __host_info__, __script_encoding__, __author__, __email__, __status__, __version__, get_isp_info())

def print_header() -> None:
    if(VERBOSE):
        print(__header__)

MAX_EXCEL_TAB_NODATE=22  # '_12212019' = 9 chars
MAX_EXCEL_TAB_DIR=27
MAX_EXCEL_TAB=31
TEMP_TAG = '~'
ALPHABET=string.ascii_uppercase

valid_chars = "-_.()~%s%s" % (string.ascii_letters, string.digits)

def sanitize_input(input_str):
    sanitized = input_str
    invalid_chars = "{}[]()<>^#+*?$@&%$!,:;/\\ "  #keep - and .
    sanitized = sanitized.replace("-","")    
    for char in invalid_chars:
        if char in input_str:
            sanitized = sanitized.replace(char,"")
    return sanitized


def generate_date_str() -> tuple:
    now = datetime.datetime.now()
    date = now.strftime("%m-%d-%Y")
    time = now.strftime("%H%M%p").lower()
    return (date, time)
    

def get_header_column_widths(input_tag_list: list) -> dict:
    # list of lists: [row1:[hdr1, hdr2, hdr3, ...], row2:[hdr1, hdr2, hdr3, ...]...]
    hdr_col_width_dict = OrderedDict([(hdr,-1) for hdr in input_tag_list[0]])
    for row_num, tag_list in enumerate(input_tag_list):
        for col_num, cell_val in enumerate(tag_list):
            alpha = ALPHABET[col_num]
            header = headers[col_num]
            max_length = len(cell_val)
            if (hdr_col_width_dict[header] < max_length):
                hdr_col_width_dict[header] = max_length+2 # each char=1, +2 for readability
    if (VERBOSE):
        print("\nDynamically sized columns widths:")
        for key, value in hdr_col_width_dict.items():
            print("   {:28} \t {} chars".format(key, value))
    return hdr_col_width_dict


def export_to_excel(output_path: str, output_filename: str, tab_name: str, stat_list_of_lists: list) -> str:
    def_name = sys._getframe().f_code.co_name.upper()
    status_str = "{}() in: '{}'\n".format(def_name, output_path)
    try:
        output_filepath = os.path.join(output_path, output_filename)
        workbook = xlsxwriter.Workbook("{}".format(output_filepath))
        worksheet1 = workbook.add_worksheet(tab_name[:MAX_EXCEL_TAB])  # <= 31 chars
        worksheet1.freeze_panes(1, 0)
        
        # https://xlsxwriter.readthedocs.io/example_conditional_format.html#ex-cond-format
        # add formating: light red fill with dark red text
        format_red = workbook.add_format({'bg_color':'#FFC7CE','font_color':'#9C0006', 'bold': False})
        # add formating: green fill with dark green text
        format_green = workbook.add_format({'bg_color':'#C6EFCE','font_color':'#006100', 'bold': False})

        hdr_col_width_dict = get_header_column_widths(stat_list_of_lists)  # includes both header and cell values in calcuation
        header_format = workbook.add_format({'bold': True, 'underline': True,'font_color':'blue', 'center_across':True })
        stat_list_of_lists.pop(0) # ignore header row
        for idx, key_hdr in enumerate(hdr_col_width_dict):
            alpha = ALPHABET[idx]
            col_width_val = hdr_col_width_dict[key_hdr]
            #print('{}:{} width{}'.format(alpha, alpha, col_width))
            worksheet1.set_column('{}:{}'.format(alpha, alpha), col_width_val)
            worksheet1.write('{}1'.format(alpha), '{}:'.format(key_hdr), header_format)
            
        centered_cells_int = workbook.add_format()
        centered_cells_int.set_num_format('0')
        centered_cells_int.set_align('center')
        centered_cells_int.set_align('vcenter')
        
        centered_cells = workbook.add_format()
        centered_cells.set_align('center')
        centered_cells.set_align('vcenter')

        date_cell_format = workbook.add_format()
        date_cell_format.set_num_format('mm/dd/yy hh:mm AM/PM')
        date_cell_format.set_align('center')
        date_cell_format.set_align('vcenter')
        
        left_centered_cells = workbook.add_format()
        left_centered_cells.set_align('left')
        left_centered_cells.set_align('vcenter')

        last_alpha = ALPHABET[len(headers)-1]  # last index -1 of len()
        worksheet1.autofilter('A1:{}65536'.format(last_alpha))

        # A0, B1, C2, D3, E4, F5, G6, H7 = 8 indexed entries
        row_num = 1
        if (len(stat_list_of_lists) > 0):
            for idx, tag_list in enumerate(stat_list_of_lists):
                row_num += 1  # not header row
                for count, tag_val in enumerate(tag_list):
                    alpha = ALPHABET[count]
                    #print('row_num: {} {} {} idx:{} count:{} tag_val:{}'.format(row_num, alpha, alpha_col_width_dict[alpha], idx, count, str(tag_val)))
                    worksheet1.write('{}{}'.format(alpha, row_num), str(tag_val), centered_cells)
                
        workbook.close()
        status =  "\nSUCCESS! {}() \n{}".format(def_name, output_filepath)
        return status
        
    except (Exception) as exception:
        status = "~!ERROR!~ {}() {}\n{}".format(def_name, sys.exc_info()[0], exception)
        return status


'''
dcmtk_headers=["filename","AccessionNumber","Modality","SourceApplicationEntityTitle","StationName","institutionName","Manufacturer","ManufacturerModelName"]
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
def build_fuji_tag_dict(input_filename: str='default') -> dict:
    def_name = sys._getframe().f_code.co_name.upper()
    fuji_tag_dict = OrderedDict()
    fuji_tag_dict['filename']=input_filename
    fuji_tag_dict['accessionNumber']='0008 0050'
    fuji_tag_dict['modality']='0008 0060'
    fuji_tag_dict['sourceApplicationEntityTitle']='0002 0016'
    fuji_tag_dict['stationName']='0008 1010'
    fuji_tag_dict['institutionName']='0008 0080'
    fuji_tag_dict['manufacturer']='0008 0070'
    fuji_tag_dict['manufacturerModelName']='0008 1090'
    return fuji_tag_dict

# tag: (0008,0050) is represented as '(0008,0050)' for DCMTK sourced files
def build_dcmtk_tag_dict(input_filename: str='default') -> dict:
    def_name = sys._getframe().f_code.co_name.upper()
    dcmtk_tag_dict = OrderedDict()
    dcmtk_tag_dict['filename']=input_filename
    dcmtk_tag_dict['accessionNumber']='(0008,0050)'
    dcmtk_tag_dict['modality']='(0008,0060)'
    dcmtk_tag_dict['sourceApplicationEntityTitle']='(0002,0016)'
    dcmtk_tag_dict['stationName']='(0008,1010)'
    dcmtk_tag_dict['institutionName']='(0008,0080)'
    dcmtk_tag_dict['manufacturer']='(0008,0070)'
    dcmtk_tag_dict['manufacturerModelName']='(0008,1090)'
    return dcmtk_tag_dict


def is_fuji_tag_dump(input_list: list) -> tuple:
    def_name = sys._getframe().f_code.co_name.upper()
    status_str = "{}() in: '{}'\n".format(def_name, os.getcwd())
    bar_counter = 0
    num_sign_counter = 0
    for line_str in input_list:
        if '|' in line_str:
            bar_counter += 1
        if '#' in line_str:
            num_sign_counter += 1
    if (num_sign_counter < bar_counter):
        isFuji = True
        isDCMTK = False
    else:
        isFuji = False
        isDCMTK = True
    if (DEBUG):
        print("   isFuji:{} bar | count:{} \n   isDCMTK: {} num # count:{}".format(isFuji, bar_counter, isDCMTK, num_sign_counter))
    return (isFuji, isDCMTK)
    

def get_tag_line_number(tag_keyword: str='(0008,0020)', input_lines: list=[]) -> int:
    # optimization: stop iterating once index to tag_keyword is located
    # using tag_keyword '(0008,0050)' or '0008 0050'
    return next((line_num for line_num, string in enumerate(input_lines) if tag_keyword in string), -1)
    # if tag_keyword not in input_lines, return -1.


def get_tag_indices(tag_keywords: list=[], input_lines: list=[], isOptimized: bool=True) -> dict:
    tag_indices_dict = OrderedDict([(hdr,'') for hdr in tag_keywords])
    # using tag_keyword '(0008,0050)' or '0008 0050'
    if (isOptimized):
        for tag_keyword in tag_keywords:
            line_num = get_tag_line_number(tag_keyword, input_lines)
            if (line_num != -1): 
                tag_indices_dict[tag_keyword]=input_lines[line_num] 
                #print("line_{:02} {:42}".format(line_num, tag_keyword))
    else:
        # interate through all lines in list to find desired tags
        for line_num, line_str in enumerate(input_lines):
            for tag_keyword in tag_keywords:
                if tag_keyword in line_str:
                    tag_indices_dict[tag_keyword]=line_str
                    #print("line_{:02} tag: {:42}".format(line_num, tag_keyword))
    if (DEBUG):
        for tag_key, tag_value in tag_indices_dict.items():
            print("{}={}".format(tag_key, tag_value))
    return tag_indices_dict


def parse_dicom_tag_dump(input_headers: list, input_folder: str='default') -> list:
    def_name = sys._getframe().f_code.co_name.upper()
    status_str = "{}() in: '{}'\n".format(def_name, os.getcwd())
    tag_name = 'tagdump'
    
    source_ext = '.txt'
    file_list = sorted(glob.glob("{}{}*{}".format(input_folder, os.sep, source_ext)))
    folder_name = os.getcwd().split('\\')[-1]
    delimiter = ','
    file_count = 0
    output_tag_list = [input_headers]  #first row contains headers
   
    if not file_list:
        error_msg = "~!ERROR!~ missing files, check directory structure..."
        print(error_msg)
    else:
        print ("PARSING: ({}) '{}' files".format(len(file_list), source_ext))
        for this_file in file_list:
            line_num = 0
            file_count += 1
            row_str = ''
            parsed_file_list = []
            read_file_handle = open(this_file, 'r')
            
            if tag_name not in this_file:
                print ("   reading_{:03}: {}".format(file_count, this_file))
                lines_list = read_file_handle.readlines()
                
                # dynamically determine which input file format: based on char counts of '|' and '#'
                isFuji, isDCMTK = is_fuji_tag_dump(lines_list)
                if (isFuji):
                    target_element_dict = build_fuji_tag_dict(this_file)
                else:
                    target_element_dict = build_dcmtk_tag_dict(this_file)
                    
                # re-initializes output dict for each file to blank values
                tag_dict = OrderedDict([(hdr,'') for hdr in list(target_element_dict.keys())])
                tag_dict['filename'] = this_file
                
                # using values: tag '(0008,0050)' or '0008 0050'
                tag_indices = get_tag_indices(list(target_element_dict.values()), lines_list)
                tag_num = 0
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
                        print("tag_{:02} {:24} \t{}  line: {:40} len:{:02} chars".format(tag_num, tag_key, tag_value, line_str, len(line_str)))
                    
                for tag_attribute, parsed_val in tag_dict.items():
                    parsed_file_list.append(parsed_val)
                    
            read_file_handle.close()
            output_tag_list.append(parsed_file_list)
            
        print("EXTRACTION: ({}) '{}' files complete".format(file_count, source_ext))
    return output_tag_list


if __name__ == "__main__":
    print("{} starting... ".format(SCRIPT_NAME))
    start = time.perf_counter()
    print_header()
    io_folder = 'tag_dumps'
    headers=["filename","accessionNumber","modality","sourceApplicationEntityTitle","stationName","institutionName","manufacturer","manufacturerModelName"]
    all_tag_list = parse_dicom_tag_dump(headers, io_folder)
    
    curr_date, curr_time = generate_date_str()
    base_name = "{}dicom_{}_{}".format(TEMP_TAG, io_folder, curr_date)
    output_filename = (base_name + '.xlsx')
    
    #works on both Linux and Windows
    output_folderpath = os.path.join(BASE_DIR, io_folder)
    xls_status = export_to_excel(output_folderpath, output_filename, base_name, all_tag_list)
    print("{}".format(xls_status))
 
    end = time.perf_counter() - start
    print("{0} finished in {1:0.2f} seconds".format(SCRIPT_NAME, end))
