#!python3
# -*- coding: utf-8 -*-
import chardet
from collections import OrderedDict
import datetime
import errno
import inspect
import hashlib
import os
import pkg_resources
import sys
from custom_lib import config

BASE_DIR, SCRIPT_NAME = os.path.split(os.path.abspath(__file__))
PARENT_PATH, CURR_DIR = os.path.split(BASE_DIR)

ASCII_CTRL = r'\/:*?"<>|'
INVALID_CHARS = ',;{}^#?$@%$'
AVOID_DIRS = ['log_files', 'Program Files', 'ProgramData', 'Windows', '.git',
              '.idea', 'venv', '__pycache__', '$RECYCLE.BIN']


def print_current_packages():
    """Helper to check which python modules are currently installed"""
    config.show_method(inspect.currentframe().f_code.co_name)
    distros = [d for d in pkg_resources.working_set]
    installed = sorted([f"{d.project_name:24}\t{d.version}" for d in distros])
    for pkg_num, pkg_name in enumerate(installed):
        print(f"pkg_{pkg_num:03d}:\t{pkg_name}")


def is_pathname_valid(path_name: str = 'default') -> bool:
    """Verify if input path is valid"""
    ERROR_INVALID_NAME = 123
    try:
        if not isinstance(path_name, str) or not path_name:
            return False
        base_name, path_name = os.path.splitdrive(path_name)
        root_dirname = os.environ.get('HOMEDRIVE', 'C:') \
            if sys.platform == 'win32' else os.path.sep
        assert os.path.isdir(root_dirname)
        root_dirname = root_dirname.rstrip(os.path.sep) + os.path.sep
        for pathname_part in path_name.split(os.path.sep):
            try:
                os.lstat(root_dirname + pathname_part)
            except OSError as exception:
                if hasattr(exception, 'winerror'):
                    if exception.winerror == ERROR_INVALID_NAME:
                        return False
                elif exception.errno in {errno.ENAMETOOLONG, errno.ERANGE}:
                    return False
    except TypeError as exception:
        print(f"  {sys.exc_info()[0]}\n{exception}")
        return False
    else:
        return True


def is_path_creatable(path_name: str = 'default') -> bool:
    """Verify if input path can be written"""
    dir_name = os.path.dirname(path_name) or os.getcwd()
    return os.access(dir_name, os.W_OK)


def is_path_exists_or_creatable(path_name: str = 'default') -> bool:
    """Verify if input path exists and can be written to"""
    try:
        return (is_pathname_valid(path_name) and
                (os.path.exists(path_name) or is_path_creatable(path_name)))
    except OSError:
        return False


def bytes_to_readable(input_bytes: int) -> str:
    """Converts file size to windows/linux formatted string"""
    number_of_bytes = float(input_bytes)
    unit_str = 'bytes'
    if config.IS_WINDOWS:
        kilobyte = 1024.0  # iso_binary size windows
    else:
        kilobyte = 1000.0  # si_unit size linux
    if input_bytes < 0:
        print(f"ERROR: input:'{number_of_bytes}' < 0")
    else:
        if (number_of_bytes / kilobyte) >= 1:
            number_of_bytes /= kilobyte
            unit_str = 'KiB'
        if (number_of_bytes / kilobyte) >= 1:
            number_of_bytes /= kilobyte
            unit_str = 'MiB'
        if (number_of_bytes / kilobyte) >= 1:
            number_of_bytes /= kilobyte
            unit_str = 'GiB'
        if (number_of_bytes / kilobyte) >= 1:
            number_of_bytes /= kilobyte
            unit_str = 'TiB'
        precision = 2
        number_of_bytes = round(number_of_bytes, precision)
    return str(f"{number_of_bytes:05.2F} {unit_str}")


def make_unicode(input_data):
    """Decodes input data into UTF8"""
    unicode_range = ('4E00', '9FFF')
    try:
        if type(input_data) != unicode_range:
            input_data = input_data.decode('UTF-8')
            return input_data
        else:
            return input_data
    except Exception as exception:
        print(f"  {sys.exc_info()[0]}\n {exception}")


def is_encoded(data, encoding: str = 'default') -> bool:
    """Verifies if bytes data is non-ASCII encoded"""
    try:
        data.decode(encoding)
        # ASCII:  7 bits 128 code points
        # Latin1: ISO-8859-1: 1-byte per char 256 code points
        # UTF-8:  1-byte to encode each char
        # UTF-16: 2-bytes to encode each char
        # cp1252: Windows codec 1-byte per char
    except UnicodeDecodeError:
        return False
    else:
        if config.DEBUG:
            print(f"{encoding} {data} {type(data)}")
        return True


def check_encoding(input_val: bytes):
    """Verifies if bytes is UTF-8"""
    if isinstance(input_val, (bytes, bytearray)):
        # bytes: immutable, bytesarray: mutable both ASCII:[ints 0<=x<256]
        return chardet.detect(input_val), input_val
    else:
        # ignore, replace backslashreplace namereplace
        bytes_arr = input_val.encode(encoding='UTF-8', errors='namereplace')
        return chardet.detect(bytes_arr), bytes_arr


def remove_accents(byte_input, byte_enc: str, confidence: float) -> str:
    """Decodes encoded bytes input data - dynamically determined"""
    dec_str = ""
    try:
        if sys.platform == 'win32':
            if confidence > 0.70:
                is_encoded(byte_input, byte_enc)
                dec_str = byte_input.decode(byte_enc)
            elif is_encoded(byte_input, 'UTF-8'):
                # input = u'û è ï - ö î ó ‘ é  í ’ ° æ ™'
                dec_str = byte_input.decode('UTF-8')
            enc_bytes = dec_str.encode('UTF-8')  # python ascii string
            if config.DEBUG:
                print(f"   byte_input:  {type(byte_input)} \t '{byte_input}'")
                print(f"   dec_str: {type(dec_str)} \t '{dec_str}'")
                print(f"   enc_bytes: {type(enc_bytes)} \t '{enc_bytes}'")
    except Exception as exception:
        print(f"\nERROR: input: '{input}' {type(input)}")
        print(f"  {sys.exc_info()[0]}\n{exception}")
    return dec_str


def get_sha1_hash(input_path: str = 'default') -> str:
    """Returns SHA1 hash hex value of input filepath"""
    fp = open(input_path, 'rb')
    rfp = fp.read()
    sha1_hash = hashlib.sha1(rfp)
    sha1_hex = str(sha1_hash.hexdigest().upper())
    fp.close()
    return sha1_hex


def get_directory_size(input_path: str = 'default') -> int:
    """Returns recursive integer sum of file sizes in input directory path"""
    config.show_method(inspect.currentframe().f_code.co_name)
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(input_path):
        for file in filenames:
            total_size += os.path.getsize(os.path.join(dirpath, file))
    return total_size


def split_path(input_path: str = 'default') -> tuple:
    """Custom os.path.split() to account for files without extensions"""
    if input_path:
        input_path_isfile = True
        # if present remove trailing os.sep '//'
        if input_path[-1] == os.sep:
            input_path = input_path[:-1]
        case_dict = {1: False, 2: False, 3: False}
        path_head, path_tail = os.path.split(input_path)
        # input_path does not include an filename or extension
        if os.path.isdir(os.path.join(path_head, path_tail)):
            input_path_isfile = False
            drive_parent_dir = str(path_head)
            curr_dir = str(path_tail)
            file_basename = ''
            file_ext = ''
        if input_path_isfile:
            full_dirpath = os.path.dirname(input_path)
            filename_wext = os.path.basename(input_path)
            file_basename, file_ext = os.path.splitext(filename_wext)
            # only split the directory (not the filename)
            sep_split_list = full_dirpath.split(os.sep)
            path_len = len(sep_split_list)
            # CASE1 where files are in drivedir
            if path_len == 1:
                case_dict[1] = True
                drive_parent_dir = str(sep_split_list[0])
                curr_dir = ''
            # CASE2 where files are only 1 subfolder from drivedir
            if path_len == 2:
                case_dict[2] = True
                drive_parent_dir = str(sep_split_list[0])
                curr_dir = str(sep_split_list[1])
            # CASE3 where files are only 2+ subfolders from drivedir
            if path_len > 2:
                case_dict[3] = True
                drive_parent_dir = os.sep.join(sep_split_list[:-1])
                curr_dir = str(sep_split_list[-1])  # last subfolder
        return drive_parent_dir, curr_dir, file_basename, file_ext


def check_pathname_skip(input_path: str = 'default') -> bool:
    """Returns true if input path is not in any directory to avoid"""
    if input_path:
        path_hit = next((s for s in AVOID_DIRS if s in input_path), True)
        if path_hit:
            return True
        else:
            return False


def sanitize(filename: str = 'default') -> str:
    """strips invalid characters from filename"""
    sanitized = filename
    for char in ASCII_CTRL:
        if char in sanitized:
            sanitized = sanitized.replace(char, "")
    for char in INVALID_CHARS:
        if char in sanitized:
            sanitized = sanitized.replace(char, "")
    return sanitized


def generate_date_str() -> tuple:
    """Creates string based on current timestamp"""
    now = datetime.datetime.now()
    date = now.strftime("%m-%d-%Y")
    time = now.strftime("%H%M%p").lower()
    return date, time


def save_output_txt(txt_path: str, input_filename: str, output_str: str,
                    delim_tag: bool = False) -> str:
    """Exports string to output text file"""
    def_name = inspect.currentframe().f_code.co_name
    try:
        if len(output_str) > 4:
            file_ext = '.log'
            if file_ext in input_filename:
                input_filename = input_filename[:-4]  # remove '.txt' extension
            if delim_tag:
                tagged_filename = f"~{input_filename}{file_ext}"
                output_path = os.path.join(txt_path, tagged_filename)
            else:
                untagged_filename = f"{input_filename}{file_ext}"
                output_path = os.path.join(txt_path, untagged_filename)
            if not os.path.exists(txt_path):
                os.makedirs(txt_path)
            # 'w'=write, 'a'=append, 'b'=binary, 'x'=create
            open_flags = 'w'
            with open(output_path, open_flags) as text_file:
                text_file.write(output_str)
            text_file.close()
            return f"\nSUCCESS! {def_name}()"
    except (IOError, OSError, PermissionError, FileExistsError) as exception:
        err_str = f"\n~!ERROR!~ {sys.exc_info()[0]}\n{exception}"
        return err_str


def count_files(input_path: str, file_ext: str = '.mp3') -> int:
    """Returns recursive count of files with specific extension"""
    file_list = []
    if is_path_exists_or_creatable(input_path):
        for (root_path, dirs, files) in os.walk(input_path):
            for this_file in files:
                if this_file.lower().endswith(file_ext):
                    file_list.append(this_file)
    return len(file_list)


def convert_dict_to_str(input_path: str, ext_map: dict) -> str:
    """Returns count of each file extension in file_path - sorted order"""
    output_str = ""
    file_count = ""
    for this_ext, this_count in ext_map.items():
        print(f"   extension_dict: [{this_ext:5} = {this_count:04}]")
        file_count += f"\t[{this_count:04}]\t{this_ext:5}\n"
    output_str = (f"FOUND: '{len(ext_map):02}' extensions: "
                  f"'{input_path}'\n{file_count}")
    return output_str


def get_extensions(input_path: str = 'default') -> dict:
    """Return recursive set of all file extensions within input path"""
    def_name = inspect.currentframe().f_code.co_name
    output_str = f"\n{def_name}() in: '{input_path}'"
    print(output_str)
    extension_dict = {}
    # get unique list of file extensions (including sub-folders)
    if is_path_exists_or_creatable(input_path):
        for root_path, dirs, files in os.walk(input_path):
            for this_file in files:
                this_file_path = os.path.join(root_path, this_file)
                if check_pathname_skip(this_file_path):
                    file_ext = f".{this_file.split('.')[-1].lower()}"
                    # exclude files without extension
                    if len(file_ext) < 8:
                        if file_ext not in extension_dict:
                            extension_dict[file_ext] = 0
    # get count of each extension in file_path - sorted order
    sorted_extension_dict = OrderedDict()
    for this_ext in sorted(extension_dict.keys()):
        file_count = count_files(input_path, this_ext)
        sorted_extension_dict[this_ext] = file_count
    return sorted_extension_dict


def get_directories(input_path: str = 'default') -> tuple:
    """Return list of directories within input path (including subfolders)"""
    def_name = inspect.currentframe().f_code.co_name
    dir_list = []
    dir_size_list_of_lists = []
    output_str = f"\n{def_name}() in: '{input_path}'\n"
    parent_dir, curr_dir, file_name, file_ext = split_path(input_path)
    par_size = get_directory_size(input_path)
    for file_dir in sorted(os.listdir(input_path)):
        if os.path.isdir(os.path.join(input_path, file_dir)):
            if file_dir not in AVOID_DIRS:
                dir_list.append(file_dir)
    output_str += (f"   ({len(dir_list)}) directories "
                   f"[{bytes_to_readable(par_size)}]\n")
    print(output_str, end='')   # skip newline to stdout
    for count, directory in enumerate(sorted(dir_list)):
        subdir_path = os.path.join(parent_dir, curr_dir, directory)
        dir_size = get_directory_size(subdir_path)
        last_mod_ts = os.path.getmtime(subdir_path)
        subdir_last_modified = datetime.datetime.fromtimestamp(last_mod_ts)
        dir_size_list = []
        dir_size_list.append(f"{count + 1:02}")
        dir_size_list.append(f"{dir_size:08}")
        dir_size_list.append(f"{bytes_to_readable(dir_size)}")
        dir_size_list.append(f"{subdir_path}")
        dir_size_list.append(f"{subdir_last_modified}")
        dir_size_list_of_lists.append(dir_size_list)
    return dir_size_list_of_lists, output_str


def get_files(input_path: str, input_ext: str) -> tuple:
    """Get files, paths for specific file extension (including sub-folders)"""
    def_name = inspect.currentframe().f_code.co_name
    file_list = []
    file_path_list = []
    if is_path_exists_or_creatable(input_path):
        for root_path, dirs, files in os.walk(input_path):
            for this_file in files:
                if this_file.lower().endswith((input_ext)):
                    this_file_path = os.path.join(root_path, this_file)
                    if check_pathname_skip(this_file_path):
                        file_list.append(this_file)
                        file_path_list.append(this_file_path)
        file_path_list.sort(key=lambda x: (-x.count(os.sep), x))
    print(f"{def_name}: [{len(file_list):03} {input_ext}] files")
    return sorted(file_list), file_path_list
