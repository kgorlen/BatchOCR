'''
Created on Jun 12, 2024

@author: Keith kgorlen@gmail.com

Batch convert PDF files to searchable PDF files using ABBYY FineReader Sprint

References:
    https://help.abbyy.com/en-us/finereader/15/user_guide/commandline_save/
    https://archive.org/details/abbyy-fine-reader-9
    https://levelup.gitconnected.com/4-python-libraries-to-convert-pdf-to-images-7a09eba83a09
    https://chatgpt.com/share/673e7bba-f57c-800d-b3fa-80f86211c094
'''

__author__ = 'Keith Gorlen kgorlen@gmail.com'

import sys
import os
import logging
from logging import Logger
from time import sleep
from datetime import datetime
import argparse
import shlex
from pathlib import Path
import glob
import re
import tempfile
import subprocess
import platform
from collections import defaultdict
from typing import NamedTuple, Optional, Generator, NoReturn

SCRIPT_DIR: Path = Path(__file__).absolute().parent
"""Path to directory containing this Python script."""
sys.path.append(str(SCRIPT_DIR))
"""Allow imports from script directory."""

from __init__ import __version__  # pylint: disable=no-name-in-module
from platformdirs import user_log_dir
import psutil
import pymupdf
from PIL import Image
import win32file

# Global Constants

SCRIPT_NAME = Path(__file__).stem
"""The current script name without .py extension."""

DATE_FMT = '%Y-%m-%d %H:%M:%S'
"""Format for dates in messages."""

STEM_SUFFIX = '_OCR_'
"""Suffix appended to stem of searchable pdf files."""

ABBYY_PATH = Path(r'C:\Program Files (x86)\ABBYY FineReader 9.0 Sprint')
"""Path to ABBYY FineReader executables."""

# Global Variables

LOGGER: Logger
"""LOGGER configured in main()."""


class ParsedArgs(NamedTuple):
    """Parsed command line options and arguments."""

    analyze: bool
    """Analyze PDF pages and report without conversion."""
    commit: bool
    """Replace original .pdf files with {STEM_SUFFIX}.pdf files."""
    debug: bool
    """Log debugging information."""
    dpi: bool
    """Resolution of PDF to TIFF conversion."""
    force: bool
    """Convert pdf file even if searchable."""
    image: float
    """Page unsearchable if image area exceeds this percent."""
    pages: int
    """Maximum number of pages to test for searchable words."""
    quiet: bool
    """Do not print INFO, WARNING, ERROR, or CRITICAL messages to `stderr`;
    default --no-quiet."""
    rollback: bool
    """Delete *{STEM_SUFFIX}.pdf files."""
    text: float
    """Page unsearchable unless text area exceeds this percent."""
    unsearchable: int
    """Convert PDF if number of unsearchable pages greater than PAGES."""
    version: bool
    """Display the version number and exit."""
    words: int
    """Minimum number of words on page for searchable test."""
    infiles: list[str]
    """Input files."""


ARGS: ParsedArgs
"""Arguments parsed by argparse() in main()."""


def info_msg(msg: str) -> None:
    """Log and print an INFO message to stdout."""
    LOGGER.info(msg)
    if not ARGS.quiet:
        print(f'{datetime.now().strftime(DATE_FMT)} - INFO - {msg}', file=sys.stderr)


def warning_msg(msg: str) -> None:
    """Log a WARNING message."""
    LOGGER.warning(msg)
    if not ARGS.quiet:
        print(f'{datetime.now().strftime(DATE_FMT)} - WARNING - {msg}', file=sys.stderr)


def error_msg(msg: str) -> None:
    """Log an ERROR message."""
    LOGGER.error(msg)
    if not ARGS.quiet:
        print(f'{datetime.now().strftime(DATE_FMT)} - ERROR - {msg}', file=sys.stderr)


def fatal_error(msg: str) -> NoReturn:
    """Log a CRITICAL message and sys.exit(1)."""
    LOGGER.critical(f'{msg}; exiting.')
    if not ARGS.quiet:
        print(f'{datetime.now().strftime(DATE_FMT)} - CRITICAL - {msg}; exiting.', file=sys.stderr)
    sys.exit(1)


def parse_command_line() -> ParsedArgs:
    """Parsed command line arguments."""

    parser = argparse.ArgumentParser(description='Batch convert PDF files to searchable PDF files')
    parser.add_argument(
        '-d',
        '--debug',
        action=argparse.BooleanOptionalAction,
        default=False,
        help='Log debug info; default False',
    )
    parser.add_argument(
        '--dpi',
        metavar='DPI',
        type=int,
        default=300,
        help='Resolution of PDF to TIFF conversion; default 300 dpi',
    )
    parser.add_argument(
        '-f',
        '--force',
        action=argparse.BooleanOptionalAction,
        default=False,
        help='Convert pdf file even if searchable; default False',
    )
    parser.add_argument(
        '--image',
        metavar='PERCENT',
        type=float,
        default=100.0,
        help='Page unsearchable if image area exceeds this percent; default 100 percent',
    )
    parser.add_argument(
        '--pages',
        metavar='PAGES',
        type=int,
        default=10,
        help='Maximum number of pages to analyze for searchability; default 10',
    )
    parser.add_argument(
        '-q',
        '--quiet',
        action=argparse.BooleanOptionalAction,
        default=False,
        help='Do not print messages; default False',
    )
    parser.add_argument(
        '--text',
        metavar='PERCENT',
        type=float,
        default=1.0,
        help='Page unsearchable unless text area exceeds this percent; default 1 percent',
    )
    parser.add_argument(
        '--unsearchable',
        metavar='PAGES',
        type=int,
        default=2,
        help='Convert PDF if number of unsearchable pages greater than PAGES; default 2',
    )
    parser.add_argument(
        '--words',
        metavar='WORDS',
        type=int,
        default=3,
        help='Minimum number of words in text box for inclusion in text area; default 3',
    )
    exclusive = parser.add_mutually_exclusive_group(required=False)
    exclusive.add_argument(
        '-a',
        '--analyze',
        action=argparse.BooleanOptionalAction,
        default=False,
        help='Analyze PDF pages and report without conversion; default False',
    )
    exclusive.add_argument(
        '--commit',
        action=argparse.BooleanOptionalAction,
        default=False,
        help=f'Replace original .pdf files with searchable *{STEM_SUFFIX}.pdf files; default False',
    )
    exclusive.add_argument(
        '--rollback',
        action=argparse.BooleanOptionalAction,
        default=False,
        help=f'Delete *{STEM_SUFFIX}.pdf files; default False',
    )
    either = parser.add_mutually_exclusive_group(required=True)
    either.add_argument(
        '-v',
        '--version',
        action=argparse.BooleanOptionalAction,
        default=False,
        help='Display the version number and exit',
    )
    either.add_argument(
        'infiles',
        metavar='FILES',
        type=str,
        nargs='*',
        default='*.pdf',
        help='List of .pdf files; input is read from stdin if "-"',
    )

    arg_list = None if len(sys.argv) > 2 else shlex.split(sys.argv[1])
    args_dict: dict[str, Optional[bool | str | list[str] | int]] = vars(parser.parse_args(arg_list))
    return ParsedArgs(**args_dict)  # type: ignore


def file_path_generator(
    globs: list[str], pattern: Optional[str] = "*", recursive: bool = True
) -> Generator[Path, None, None]:
    """Generate matched files from a list of files and/or directories with
    globbing, environment variable substitution, '~' expansion, '-' input
    from stdin, and optional '**' recursion.

    Parameters
    ----------
    globs : list[str]
        List of file and directory paths to process.
    pattern : Optional[str], optional
        Glob pattern to match files, by default "*"
    recursive : bool, optional
        Process "**" directories recursively, by default True

    Yields
    ------
    Generator[Path, None, None]
        List of Path objects for matched files.
    """
    assert pattern is not None

    def expand_glob(glob_pattern: str) -> Generator[Path, None, None]:
        """Expand a glob

        Parameters
        ----------
        glob_pattern : str
            A glob.

        Yields
        ------
        Generator[Path, None, None]
            Paths from expanded glob.
        """

        for filename in glob.iglob(glob_pattern, recursive=recursive):
            path = Path(os.path.expanduser(os.path.expandvars(filename)))
            if path.is_file() and path.match(pattern):
                yield path

    for glob_pattern in globs:
        if glob_pattern == '-':
            LOGGER.debug('Reading FILES from stdin ...')
            for stdin_glob in sys.stdin.read().splitlines():
                yield from expand_glob(stdin_glob.strip())
            continue

        path = Path(os.path.expanduser(os.path.expandvars(glob_pattern)))
        if path.is_dir():
            yield from expand_glob(str(path.joinpath(pattern)))
        else:
            yield from expand_glob(glob_pattern)


def analyze_pdf(pdf_file: Path) -> defaultdict[str, list[int]]:
    """Check if a PDF is searchable.

    Parameters
    ----------
    file_file : Path
        Input PDF file Path.

    Returns
    -------
    defaultdict[str, list[int]]
        Lists of page numbers that are 'Blank', 'Unsearchable' or 'Searchable'.

    Notes
    -----
    https://stackoverflow.com/questions/55704218/how-to-check-if-pdf-is-scanned-image-or-contains-text
    https://stackoverflow.com/questions/63494812/how-can-i-distinguish-a-digitally-created-pdf-from-a-searchable-pdf
    https://www.petercollingridge.co.uk/blog/language/analysing-english/basic-analysis/
    """
    pdf_type: defaultdict[str, list[int]] = defaultdict(list)

    document = pymupdf.open(pdf_file)

    for pagenum, page in enumerate(document, 1):

        if pagenum > ARGS.pages:
            break

        if not (
            page.get_text().strip()
            or len(page.get_images(full=True)) > 0
            or len(page.get_drawings()) > 0
        ):
            pdf_type['Blank'].append(pagenum)
            LOGGER.debug(f'{pdf_file} page {pagenum}: Blank.')
            continue

        page_area = abs(page.rect)
        """Total page area."""
        img_area = 0.0
        """Total image area on page."""
        text_area = 0.0
        """Total area of text boxes with words on page."""
        text_count = 0
        """Number of text boxes on page."""
        text_skipped = 0
        """Number of text boxes without words on page."""

        text_dict = page.get_text("dict")
        for block in text_dict["blocks"]:
            bbox = block["bbox"]
            x0, y0, x1, y1 = bbox
            area = (x1 - x0) * (y1 - y0)

            if block["type"] == 0:  # Text block
                text_count += 1
                text = ' '.join([span["text"] for line in block["lines"] for span in line["spans"]])
                words = []
                for m in re.finditer(r'\w{5,20}', text):
                    words.append(m.group(0))
                    if len(words) == ARGS.words:
                        text_area += area
                        # LOGGER.debug(
                        #     f'{pdf_file} page {pagenum}: Words found  '
                        #     f'"{'" "'.join(words)}".'
                        # )
                        break
                else:
                    text_skipped += 1

            elif block["type"] == 1:  # Image block
                img_area += area

            else:
                LOGGER.debug(f'{pdf_file} page {pagenum}: unknown block type {block["type"]}.')

        text_pct = text_area / page_area * 100
        img_pct = img_area / page_area * 100

        if text_area > page_area:
            LOGGER.warning(
                f'{pdf_file} page {pagenum}: ' f'Text area ({text_area}) > page area ({page_area}).'
            )

        if text_area > page_area:
            LOGGER.warning(
                f'{pdf_file} page {pagenum}: ' f'Image area ({img_area}) > page area ({page_area}).'
            )

        LOGGER.debug(
            f'{pdf_file} page {pagenum}: Text area {text_area:.0f} ({text_pct:.1f}%), '
            f'{text_skipped} of {text_count} blocks skipped; '
            f'Image area {img_area:.0f} ({img_pct:.1f}%).'
        )

        if text_pct <= ARGS.text or img_pct > ARGS.image:
            pdf_type['Unsearchable'].append(pagenum)
        else:
            pdf_type['Searchable'].append(pagenum)

    LOGGER.debug(f'{pdf_file}: {", ".join([f'{k} {n}' for k, n in pdf_type.items()])}.')

    return pdf_type


def pdf_to_multipage_tiff(pdf_path, tiff_path):
    """
    Convert a PDF (RGB or grayscale) to a multipage TIFF file.

    Parameters:
        pdf_path (str): Path to the input PDF file.
        tiff_path (str): Path to the output TIFF file.
    """
    LOGGER.debug(f'Converting {pdf_path} to tiff ...')
    pdf_document = pymupdf.open(pdf_path)
    images = []

    # Iterate through each page
    for page_number, page in enumerate(pdf_document, 1):
        LOGGER.debug(f'Rendering page {page_number} ...')
        pix = page.get_pixmap(dpi=ARGS.dpi)

        # Convert to a Pillow Image
        mode = "RGB" if pix.n > 1 else "L"  # RGB or grayscale
        img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
        images.append(img)

    # Save as multipage TIFF
    if not images:
        LOGGER.debug('No images rendered.')
        return

    LOGGER.debug('Saving images to tiff file ...')
    images[0].save(tiff_path, save_all=True, append_images=images[1:], compression="tiff_deflate")

    pdf_document.close()
    tiff_path.close()


def convert_to_searchable_pdf(image_path: Path, output_pdf: Path) -> None:
    """Use ABBYY FineReader CLI to convert images to searchable PDF.

    Parameters
    ----------
    image_path : _type_
        _description_
    output_pdf : _type_
        _description_
    """
    temp_pdf: Path = output_pdf.with_stem(output_pdf.stem + 'TMP_')
    cmd = rf'{ABBYY_PATH}\finecmd.exe "{image_path}" /out "{temp_pdf}"'
    LOGGER.debug(f'Starting {cmd} ...')

    try:
        with subprocess.Popen(cmd, creationflags=subprocess.CREATE_NEW_PROCESS_GROUP) as fine:
            LOGGER.debug(f'OCR process {fine.pid} started ...')

            while not temp_pdf.exists():
                if (exit_status := fine.poll()) is not None:
                    fatal_error(f'{cmd}: exited with status {exit_status}')

                sleep(1)

            LOGGER.debug(f'OCR process {fine.pid} running ...')

            # https://stackoverflow.com/questions/66158631/check-if-a-file-is-written-or-in-use-by-another-process
            while True:
                try:
                    temp_pdf.rename(output_pdf)
                    break
                except (PermissionError, OSError) as e:
                    LOGGER.debug(f'{temp_pdf}.rename({output_pdf} failed: {e}, retrying ...)')
                    if (exit_status := fine.poll()) is not None:
                        fatal_error(f'{cmd}: exited with status {exit_status}')
                sleep(1)

            LOGGER.debug('OCR process done, terminating ...')
            fine.terminate()

            try:
                status = fine.wait(timeout=10)
                LOGGER.debug(f'FineCmd terminated with exit status {status}')
            except subprocess.TimeoutExpired:
                fatal_error(f'{cmd}: Failed to terminate FineCmd.exe')

        this_proc = psutil.Process()
        username = this_proc.username()
        LOGGER.debug('Checking for Sprint process ...')
        for p in psutil.process_iter(attrs=['pid', 'name', 'username', 'cmdline']):
            if p.info['name'] == 'Sprint.exe' and p.info['username'] == username:
                LOGGER.debug(f'Found: Sprint {p.info['cmdline']}')
                break
        else:
            LOGGER.debug('Sprint process not running.')
            return

        LOGGER.debug('Terminating Sprint.exe ...')
        p.terminate()

        try:
            LOGGER.debug('Waiting for Sprint.exe to terminate ...')
            status = p.wait(timeout=10)
        except subprocess.TimeoutExpired:
            fatal_error(f'{cmd}: Failed to terminate Sprint.exe')

        LOGGER.debug(f'Sprint status: {status}')

    except Exception as e:
        LOGGER.debug(f'Deleting {temp_pdf} ...')
        temp_pdf.unlink(missing_ok=True)
        raise e.with_traceback(e.__traceback__)


def copy_creation_time(source_file: Path, target_file: Path) -> None:
    """Copy creation time from source_path to target_path (Windows only)

    Parameters
    ----------
    source_file : Path
        Source file.
    target_file : Path
        Creation time will be set to that of the source_file.
    """
    LOGGER.debug(f'Copying creation time from {source_file} to {target_file} ...')

    # Retrieve the creation time of the source file
    handle = win32file.CreateFile(
        str(source_file),
        win32file.GENERIC_READ,
        win32file.FILE_SHARE_READ,
        None,
        win32file.OPEN_EXISTING,
        win32file.FILE_ATTRIBUTE_NORMAL,
        None,
    )
    # GetFileTime returns a tuple of FILETIME values
    creation_time = win32file.GetFileTime(handle)[0]
    win32file.CloseHandle(handle)

    # Set the creation time of the target file
    handle = win32file.CreateFile(
        str(target_file),
        win32file.GENERIC_WRITE,
        win32file.FILE_SHARE_WRITE,
        None,
        win32file.OPEN_EXISTING,
        win32file.FILE_ATTRIBUTE_NORMAL,
        None,
    )
    win32file.SetFileTime(handle, creation_time, None, None)  # Set only the creation time
    win32file.CloseHandle(handle)
    LOGGER.debug(f'Creation time of "{source_file}" ({creation_time}) copied to "{target_file}".')


def main() -> None:
    """Batch convert PDF files to searchable PDF files."""

    global ARGS, LOGGER

    # Initialize logger

    LOGGER = logging.getLogger(__name__)
    logfile: Path = (
        Path(user_log_dir('BatchOCR', appauthor=False, ensure_exists=True)) / f'{SCRIPT_NAME}.log'
    )
    logging.basicConfig(
        filename=logfile,
        filemode='a',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )

    LOGGER.info(f'===> {SCRIPT_NAME} version {__version__}')

    # Parse command line arguments

    ARGS = parse_command_line()
    LOGGER.info(f'{SCRIPT_NAME} arguments: {ARGS}.')

    if ARGS.debug:
        LOGGER.setLevel(logging.DEBUG)

    if ARGS.version:
        info_msg(f'{SCRIPT_NAME} version {__version__}')
        sys.exit(0)

    if ARGS.commit:
        # Replace original PDF files with *_OCR_.pdf files.
        for pdf_file in file_path_generator(ARGS.infiles, pattern=f'*{STEM_SUFFIX}.pdf'):
            assert re.match(rf'.+{STEM_SUFFIX}\.pdf$', pdf_file.name)
            orig_pdf = Path.joinpath(
                pdf_file.parent, re.sub(rf'(.*){STEM_SUFFIX}.pdf$', r'\1.pdf', pdf_file.name)
            )
            LOGGER.info(f'Renaming "{pdf_file}" to "{orig_pdf}" ...')
            orig_pdf.unlink(missing_ok=True)
            pdf_file.rename(orig_pdf)
            info_msg(f'Renamed "{pdf_file}" to "{orig_pdf}".')
        return

    if ARGS.rollback:
        # Delete *_OCR_.pdf files.
        for pdf_file in file_path_generator(ARGS.infiles, pattern=f'*{STEM_SUFFIX}.pdf'):
            assert re.match(rf'.+(?-i:{STEM_SUFFIX})\.pdf$', pdf_file.name, re.IGNORECASE)
            LOGGER.info(f'Deleting "{pdf_file}" ...')
            pdf_file.unlink()
            info_msg(f'Deleted "{pdf_file}".')
        return

    if ARGS.analyze:
        for pdf_file in file_path_generator(ARGS.infiles, pattern='*.pdf'):
            pdf_type = analyze_pdf(pdf_file)
            info_msg(f'{pdf_file}: {", ".join([f'{k} {n}' for k, n in pdf_type.items()])}')
        return

    for pdf_file in file_path_generator(ARGS.infiles, pattern='*.pdf'):

        # Skip *_OCR_.pdf files
        if re.fullmatch(rf'.*(?-i:{STEM_SUFFIX})\.pdf', pdf_file.name, re.IGNORECASE):
            info_msg(f'Searchable PDF SKIPPED:"{pdf_file}".')
            continue

        ocr_pdf = pdf_file.with_name(pdf_file.stem + f'{STEM_SUFFIX}.pdf')
        if ocr_pdf.exists():
            if ARGS.force:
                ocr_pdf.unlink()
            else:
                info_msg(f'Converted PDF SKIPPED: "{ocr_pdf}".')
                continue

        if ARGS.force:
            LOGGER.info(f'"{pdf_file}" --force converting ...')
        elif (len((pdf_type := analyze_pdf(pdf_file))['Searchable']) == 0
              or len(pdf_type['Unsearchable']) > ARGS.unsearchable):
            LOGGER.info(f'"{pdf_file}" is not searchable, converting ...')
        else:
            info_msg(f'Searchable PDF SKIPPED:"{pdf_file}".')
            continue

        with tempfile.NamedTemporaryFile(suffix=".tiff", delete_on_close=False) as tiff_file:

            # Convert PDF to images
            LOGGER.info(f'Writing pdf images to "{tiff_file.name}" ...')
            pdf_to_multipage_tiff(pdf_file, tiff_file)

            # Convert images to searchable PDF using ABBYY
            LOGGER.info(f'Converting "{pdf_file}" to searchable PDF "{ocr_pdf}" ...')
            convert_to_searchable_pdf(Path(tiff_file.name), ocr_pdf)
            info_msg(f'Searchable PDF CREATED: "{ocr_pdf}".')

            # Copy creation time
            if platform.system() == "Windows":
                copy_creation_time(pdf_file, ocr_pdf)


if __name__ == "__main__":
    main()
