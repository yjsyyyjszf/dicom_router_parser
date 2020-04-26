# -*- coding: UTF-8 -*-
"""File tools module to for basic file I/O utilities."""
import datetime
import inspect
import hashlib
import os
import pathlib
import sys
import chardet
from collections import OrderedDict
from boltons.fileutils import mkdir_p


BASE_DIR, SCRIPT_NAME = os.path.split(os.path.abspath(__file__))
PARENT_PATH, CURR_DIR = os.path.split(BASE_DIR)
IS_WINDOWS = sys.platform.startswith('win')
DEBUG = False
SHOW_METHODS = False

__all__ = ['show_method', 'bytes_to_readable', 'split_path', 'sanitize',
           'save_output_txt', 'get_directories', 'get_files', 'get_extensions']


def show_method(method_name: str) -> None:
    """Display method names for verbose/debugging."""
    if SHOW_METHODS:
        print(f"{method_name.upper()}()")


def bytes_to_readable(input_bytes: int) -> str:
    """Converts file size to Windows/Linux formatted string."""
    if not isinstance(input_bytes, int) or not input_bytes:
        return '0 bytes'
    number_of_bytes = float(input_bytes)
    unit_str = 'bytes'
    if IS_WINDOWS:
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


def is_encoded(data, encoding: str = 'default') -> bool:
    """Verifies if bytes object is non-ASCII encoded."""
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
        return True


def check_encoding(input_val: bytes):
    """Verifies if bytes object is UTF-8."""
    if isinstance(input_val, (bytes, bytearray)):
        # bytes: immutable, bytesarray: mutable both ASCII:[ints 0<=x<256]
        return chardet.detect(input_val), input_val
    # options: ignore, replace, backslashreplace, namereplace
    bytes_arr = input_val.encode(encoding='UTF-8', errors='namereplace')
    return chardet.detect(bytes_arr), bytes_arr


def remove_accents(byte_input, byte_enc: str, confidence: float) -> str:
    """Decodes bytes object - dynamically determined."""
    # unicode_example = u'û è ï - ö î ó ‘ é  í ’ ° æ ™'
    dec_str = ""
    try:
        if sys.platform == 'win32':
            if confidence > 0.70:
                is_encoded(byte_input, byte_enc)
                dec_str = byte_input.decode(byte_enc)
            elif is_encoded(byte_input, 'UTF-8'):
                dec_str = byte_input.decode('UTF-8')
            # enc_bytes = dec_str.encode('UTF-8')  # python ascii string
    except (UnicodeDecodeError, UnicodeError) as exception:
        print(f"\nERROR: input: '{input}' {type(input)}")
        print(f"  {sys.exc_info()[0]}\n{exception}")
    return dec_str


def get_sha256_hash(input_path: str = 'default') -> str:
    """Returns SHA1 hash value of input filepath."""
    if not isinstance(input_path, str) or not input_path:
        return 'no hash'
    pl_path = pathlib.Path(input_path)
    if pl_path.exists():
        file_pointer = open(input_path, 'rb')
        fp_read = file_pointer.read()
        sha_hash = hashlib.sha256(fp_read)
        sha_hex = str(sha_hash.hexdigest().upper())
        file_pointer.close()
        return sha_hex


def get_directory_size(input_path: str = 'default') -> int:
    """Returns recursive sum of file sizes from input directory path."""
    show_method(inspect.currentframe().f_code.co_name)
    if not isinstance(input_path, str) and not input_path:
        return 0
    pl_path = pathlib.Path(input_path)
    if pl_path.exists():
        r_file_sizes = [f.stat().st_size for f in pl_path.rglob('*')
                                if f.is_file()]
        return sum(r_file_sizes)
    return 0


def split_path(input_path: str = 'default') -> tuple:
    if isinstance(input_path, str) or input_path:
        pl_path = pathlib.Path(input_path)
        if pl_path.exists():
            path_head, path_tail = os.path.split(input_path)
            parent = pl_path.parent
            curr_dir = str(path_tail)
            file_name = pl_path.stem
            file_ext = pl_path.suffix
            return parent, curr_dir, file_name, file_ext
    return None


def path_not_in_avoid(input_path: str = 'default') -> bool:
    """Returns true if all directories to avoid are not in input_path."""
    if isinstance(input_path, str) and input_path:
        avoid_dirs = ['log_files', 'Program Files', 'ProgramData', 'Windows',
                      '.git', '.idea', 'venv', '__pycache__', '$RECYCLE.BIN',
                      '__init__']
        path_hit = next((s for s in avoid_dirs if s in input_path), 'VALID')
        if path_hit == 'VALID':
            return True
    return False


def sanitize(filename: str = 'default') -> str:
    """Strips invalid filepath characters from input string."""
    sanitized = filename
    ascii_control = r'\/:*?"<>|'
    for char in ascii_control:
        if char in sanitized:
            sanitized = sanitized.replace(char, '')
    invalid_path_chars = ',;{}^#?$@%$'
    for char in invalid_path_chars:
        if char in sanitized:
            sanitized = sanitized.replace(char, '')
    return sanitized


def generate_date_str() -> tuple:
    """Creates string based on current timestamp when called."""
    now = datetime.datetime.now()
    date = now.strftime("%m-%d-%Y")
    time = now.strftime("%H%M%p").lower()
    return date, time


def save_output_txt(out_path: str, output_filename: str, output_str: str,
                    delim_tag: bool = False,
                    replace_ext: bool = True) -> str:
    """Exports string to text file, uses the '.txt' file extension."""
    def_name = inspect.currentframe().f_code.co_name
    try:
        if len(output_str) > 4:
            # split based on last occurrence of '.' using rsplit()
            if '.' in output_filename:
                base_name, orig_ext = output_filename.rsplit(sep='.',
                                                             maxsplit=1)
            else:
                base_name, orig_ext = (output_filename, '')
            if replace_ext:
                out_filename_ext = f"{base_name}.txt"
            else:
                out_filename_ext = f"{base_name}.{orig_ext}"
            if delim_tag:
                out_filename_ext = f"~{out_filename_ext}"
            output_path_txt = os.path.join(out_path, out_filename_ext)
            if not os.path.exists(out_path):
                mkdir_p(out_path)
            # 'w'=write, 'a'=append, 'b'=binary, 'x'=create
            open_flags = 'w'
            with open(output_path_txt, open_flags) as txt_file:
                txt_file.write(output_str)
            txt_file.close()
            status = f"\nSUCCESS! {def_name}()"
        else:
            status = f"\nERROR! no data to export... {def_name}()"
    except (IOError, OSError, PermissionError, FileExistsError) as exception:
        status = f"\n~!ERROR!~ {sys.exc_info()[0]}\n{exception}"
    finally:
        return status


def count_files(input_path: str, file_ext: str = '.mp3') -> int:
    """Returns recursive count of files with specific extension."""
    if not isinstance(input_path, str) and not input_path:
        return 0
    pl_path = pathlib.Path(input_path)
    if pl_path.exists():
        if isinstance(file_ext, str) and file_ext:
            files = sorted(pl_path.rglob(f"*{file_ext}"))
            return len(files)
    return 0


def convert_dict_to_str(input_path: str, ext_map: dict) -> str:
    """Converts dictionary values to custom formatted string."""
    output_str = ""
    file_count = ""
    if ext_map:
        for _ext, count in ext_map.items():
            print(f"   extension_dict: [{_ext:5} = {count:04}]")
            file_count += f"\t[{count:04}]\t{_ext:5}\n"
        output_str = (f"FOUND: '{len(ext_map):02}' extensions: "
                      f"'{input_path}'\n{file_count}")
    return output_str


def get_extensions(input_path: str = 'default') -> dict:
    """Returns recursive set of all file extensions within input path."""
    def_name = inspect.currentframe().f_code.co_name
    output_str = f"\n{def_name}()"
    print(output_str)
    extension_dict = {}
    # get unique set of file extensions (including sub-folders)
    if pathlib.Path(input_path).exists():
        for root_path, dirs, files in os.walk(input_path):
            for _file in files:
                if path_not_in_avoid(_file):
                    file_ext = f".{_file.split('.')[-1].lower()}"
                    # exclude files without extension
                    if len(file_ext) > 1:
                        if file_ext not in extension_dict:
                            extension_dict[file_ext] = 0
    # get count of each unique extension in file_path - sorted order
    sorted_extension_dict = OrderedDict()
    for _ext in sorted(extension_dict.keys()):
        file_count = count_files(input_path, _ext)
        sorted_extension_dict[_ext] = file_count
    return sorted_extension_dict


def get_directories(input_path: str = 'default') -> tuple:
    """Return list of directories within input path (including subfolders)."""
    def_name = inspect.currentframe().f_code.co_name
    dir_size_list = []
    output_str = f"\n{def_name}()\n"
    if isinstance(input_path, str) and input_path:
        dir_list = []
        parent_dir, curr_dir, file_name, file_ext = split_path(input_path)
        par_size = get_directory_size(input_path)
        for file_dir in sorted(os.listdir(input_path)):
            if os.path.isdir(os.path.join(input_path, file_dir)):
                if path_not_in_avoid(file_dir):
                    dir_list.append(file_dir)
        output_str += (f"   ({len(dir_list)}) directories "
                       f"[{bytes_to_readable(par_size)}]\n")
        print(output_str, end='')  # skip newline to stdout
        for count, directory in enumerate(sorted(dir_list)):
            subdir_path = os.path.join(parent_dir, curr_dir, directory)
            dir_size = get_directory_size(subdir_path)
            last_mod_ts = os.path.getmtime(subdir_path)
            subdir_last_modified = datetime.datetime.fromtimestamp(last_mod_ts)
            dir_size = [f"{count + 1:02}",
                        f"{dir_size:08}",
                        f"{bytes_to_readable(dir_size)}",
                        f"{subdir_path}",
                        f"{subdir_last_modified}"]
            dir_size_list.append(dir_size)
    return dir_size_list, output_str


def get_files(input_path: str, input_ext: str) -> tuple:
    """Get files, paths for specific file extension (including sub-folders)."""
    def_name = inspect.currentframe().f_code.co_name
    file_list = []
    file_path_list = []
    if isinstance(input_path, str) and isinstance(input_ext, str):
        if pathlib.Path(input_path).exists():
            for root_path, dirs, files in os.walk(input_path):
                for _file in files:
                    if _file.lower().endswith(input_ext):
                        file_path = os.path.join(root_path, _file)
                        if path_not_in_avoid(file_path):
                            file_list.append(_file)
                            file_path_list.append(file_path)
            file_path_list.sort(key=lambda x: (-x.count(os.sep), x))
        print(f"{def_name}: [{len(file_list):03} {input_ext}] files")
    return sorted(file_list), file_path_list
