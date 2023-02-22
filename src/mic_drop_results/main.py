import contextlib
import ctypes
from io import BytesIO
import itertools
from multiprocessing import Pool, freeze_support
import os
from signal import signal, SIGINT, SIG_IGN
from subprocess import check_call, run, DEVNULL
import sys
import time
import webbrowser

from colorama import init, Fore, Style
import numpy as np
import pandas as pd
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE  # type: ignore
from pptx.enum.text import PP_ALIGN  # type: ignore
from pptx.slide import Slide
from pptx.util import Cm
from pywintypes import com_error
import requests
import win32com.client

from client import ProgramStatus, fetch_latest_version
from client import download_avatar
from compiled_regex import *
from config import Config
from constants import *
from errors import Error, ErrorType, print_exception_hook
from exceptions import *
from utils import is_number, as_type, hex_to_rgb, parse_version, abs_path
from utils import inp, disable_console, enable_console, console_style, bold
from utils import get_avatar_path, artistic_effect, parse_coef, clean_name
from utils import ProgressBar
from vba.macros import module1_bas


def _replace_avatar(slide: Slide, shape, run, *, uid: str) -> None:
    eff = parse_coef(run.text, field_name='p')
    run.text = ''

    og_path = get_avatar_path(uid)
    if not og_path.is_file():
        return None

    # Add avatar to slide
    avatar_path = artistic_effect(og_path, effect=eff)
    new_shape = slide.shapes.add_picture(  # type: ignore
        str(avatar_path), shape.left, shape.top,
        shape.width, shape.height)
    new_shape.auto_shape_type = MSO_SHAPE.OVAL
    old = shape._element
    new = new_shape._element
    old.addnext(new)
    old.getparent().remove(old)


def _replace_text(run, field_name, *, text: str) -> None:
    if text == 'nan':
        text = ''

    if field_name.startswith(cfg.trigger_word):  # conditional formatting
        for ind, seg_point in enumerate(cfg.ranges[::-1]):
            if not is_number(text):
                break

            if float(text) >= seg_point:
                if parse_coef(run.text, field_name=field_name) == 0:
                    run.font.color.rgb = RGBColor(*scheme[::-1][ind])
                else:
                    run.font.color.rgb = RGBColor(*scheme_alt[::-1][ind])
                break

        run.text = text
    else:
        run.text = run.text.replace('{'+field_name+'}', text)


def _replace_image_url(shape, p, run) -> None:
    if img_url := url_pattern.findall(run.text):
        with contextlib.suppress(Exception):
            margin_left = _insert_image(shape, img_url=img_url[0])
            run.text = run.text.replace(img_url[0], '')
            # After some experiments, I have measured that 12.47 cm = 4490850
            # Therefore, we have 1 cm = 360132.3175621492
            shape.text_frame.margin_left = Cm(margin_left/360132.3175621492)
            p.alignment = PP_ALIGN.LEFT


def _insert_image(shape, *, img_url: str) -> float:
    im_bytes = BytesIO(requests.get(img_url).content)
    img = Image.open(im_bytes)

    height = shape.height
    width = height/img.height * img.width
    left = shape.left + (shape.width-width)/2
    top = shape.top

    new_shape = slide.shapes.add_picture(  # type: ignore
        im_bytes,
        left, top, width, height
    )
    shape._element.addnext(new_shape._element)

    return left + width  # left margin for the remaining text


def fill_slide(slide: Slide, data: dict[str, str]) -> None:
    # https://wiki.python.org/moin/TimeComplexity
    cols = set([*data] + ['p'])

    for shape in slide.shapes:  # type: ignore
        if not shape.has_text_frame:
            continue

        for p in shape.text_frame.paragraphs:
            for run in p.runs:
                for field_name in field_name_pattern.findall(run.text):
                    field_name = field_name.lstrip('__')
                    if field_name not in cols:  # O(1) time complexity
                        continue

                    # Replace {p} with avatar
                    if field_name == 'p':
                        _replace_avatar(slide, shape, run,
                                        uid=data['uid'])
                        break

                    # Replace text
                    _replace_text(run, field_name,
                                  text=data[field_name])

                    _replace_image_url(shape, p, run)


def preview_df(df: pd.DataFrame, filter_series: pd.Series | None = None, *,
               n_cols: int, n_cols_ext: int = 5,
               highlight: bool = True,
               words_to_highlight: list[str | None] | None = None) -> str:
    """Formats and returns a string preview of the dataframe.

    Args:
        df: the dataframe to format.
        filter_series (optional): the boolean series used to select
            rows. Pass None to include all rows in the preview. Defaults
            to None.
        n_cols: the number columns to format (starts at first column).
            If n_cols is less than the number of columns in the dataframe,
            the preview will be a snippet, which shows ellipses at the
            end of every line).
        n_cols_ext (optional): the number of extra columns to preview
            alongside with the formatted columns. Defaults to 5.
        highlight (optional): enable value highlighting. Defaults to
            True.
        words_to_highlight (optional): list of words
            to highlight in red (only effective to values in the first
            n_cols columns). List None if the value to highlight is
            np.nan. Pass None to highlight nothing. Defaults to None.

    Returns:
        str: the formatted dataframe as a string.
    """
    if words_to_highlight is None:
        words_to_highlight = []

    df = df.copy(deep=True)
    if filter_series is not None:
        df = df[filter_series]
    df.index += 2  # reflect row index as displayed in Excel

    # Only show the first few columns for preview
    df = df.iloc[:, : min(n_cols+n_cols_ext, len(df.columns))]

    # Replace values_to_highlight with ⁅values_to_highlight⁆
    prefix, suffix = '⁅', '⁆'  # must be single length
    highlight_str = lambda x: f'{prefix}{x}{suffix}'
    for word in words_to_highlight:
        if word is None:
            df.iloc[:, :n_cols] = df.iloc[:, :n_cols].fillna(
                highlight_str('NaN'))
        else:
            df.iloc[:, :n_cols] = df.iloc[:, :n_cols].replace(
                word, highlight_str(word))

    preview = (repr(df.head(8)) if n_cols < len(df.columns)
               else repr(df))

    # Highlight values
    preview = (preview.replace(prefix, Fore.RED + '  ')
                      .replace(suffix, Fore.RESET))

    # Highlight column names
    for n in range(n_cols):
        col = df.columns.tolist()[n]
        preview = preview.replace(
            ' ' + col, f' {Fore.RED}{col}{Fore.RESET}', 1)

    # Add ... at the end of each line if preview is a snippet
    if n_cols < len(df.columns):
        preview = preview.replace('\n', '  ...\n') + '  ...'

    # Bold first row
    preview = f'{Style.BRIGHT}' + preview.replace('\n', Style.NORMAL + '\n', 1)

    if not highlight:
        preview = preview.replace(Fore.RED, '')
    return preview


def _import_avatars():
    failed = False
    max_attempt = 5
    has_task = False  # skip avatar download banner if no download task exists
    uids_unknown = []
    pool = Pool(min(4, len(token_list) + 1))

    for attempt in range(1, max_attempt+1):
        uids = []
        for df in groups.values():
            if df['__uid'].dtype.kind in 'biufc':
                Error(70).throw()

            for id in df['__uid']:
                if not (pd.isnull(id)
                        or get_avatar_path(id).is_file()
                        or id in uids
                        or id in uids_unknown):
                    uids.append(id)
        if not uids:
            break

        queue_len = len(uids)
        if attempt == 1:
            # Initialize download task
            has_task = True
            print(f'\n\nDownloading avatars... ({queue_len} in queue)')
            print('Make sure your internet connection is stable while '
                  'we are downloading.')
        elif attempt >= max_attempt:
            failed = True
            uids_unknown += uids
            break

        try:
            pool.starmap(
                download_avatar, zip(
                    uids,
                    itertools.islice(  # distribute tokens evenly
                        itertools.cycle(token_list), queue_len),
                    [cfg.avatar_resolution] * queue_len
                ))
        except (ConnectionError, TimeoutError) as e:
            if attempt >= 3: Error(20).throw(err_type=ErrorType.WARNING)
        except InvalidTokenError as e:
            Error(21.1).throw(*e.args)
        except DiscordAPIError as e:
            Error(22).throw(*e.args)

    if uids_unknown:
        Error(23).throw(str(uids_unknown), err_type=ErrorType.WARNING)

    if has_task and not failed:
        print('\033[A\033[2K\033[A\033[2K' + 'Avatar download complete!')
        pool.close()
        pool.join()



if __name__ == '__main__':
    version_tag = '3.0'
    ctypes.windll.kernel32.SetConsoleTitleW('Mic Drop Results')

# Section A: Fix console-related issues
    freeze_support()          # multiprocessing freeze support
    signal(SIGINT, SIG_IGN)   # handle KeyboardInterrupt
    disable_console()
    sys.excepthook = print_exception_hook  # avoid exiting program on exception
    init()                                 # enable ANSI escape sequences

# Section B: Check for missing files
    if missing_files := [f for f in (
            'settings.ini', 'token.txt',
            'template.pptm', 'data.xlsx',
        ) if not abs_path(f).is_file()]:
        Error(40).throw(
            'The following files are missing:',
            '- ' + '\n- '.join(missing_files),
            f'{bold("Current working directory:")}  {MAIN_DIR}')

# Section C: Load user configurations
    cfg = Config(str(abs_path('settings.ini')))
    n_scols = len(cfg.sort_orders)  # number of sorting columns
    avatar_mode = cfg.avatar_mode  # is subject to change later
    scheme, scheme_alt = [
        list(map(hex_to_rgb, x)) for x in (cfg.scheme, cfg.scheme_alt)
    ]

# Section D: Parse and test tokens
    with open(abs_path('token.txt')) as f:
        lines = f.read().splitlines()
        token_list = [line.replace('"', '').strip()
                      for line in lines if len(line) > 70]

    if avatar_mode and not token_list:
        Error(21).throw()


# Section E: Check for updates
    status = None
    if cfg.update_check:
        with contextlib.suppress(requests.exceptions.ConnectionError,
                                 requests.exceptions.ReadTimeout, KeyError):
            # Fetch the latest version and the summary of the update
            latest_tag, summary = fetch_latest_version()

            latest, current = parse_version(
                latest_tag, version_tag
            )

            if latest > current:
                status = ProgramStatus.UPDATE_AVAILABLE

                console_style(Fore.YELLOW, Style.BRIGHT)
                print(f'Update v{latest_tag}')

                console_style(Style.NORMAL)
                print(summary)

                print(LATEST_RELEASE_URL)
                print()
                console_style()

                webbrowser.open(LATEST_RELEASE_URL, new=2)

            elif latest < current:
                status = ProgramStatus.BETA
            else:
                status = ProgramStatus.UP_TO_DATE

    # Print a header containing information about the program

    # Normal:       Mic Drop Results (v3.10) [latest]
    #               https://github.com/banz04/mic-drop-results


    # With update:  Update v3.11
    #               A summary of the update will appear in this line.
    #               https://github.com/banz04/mic-drop-results/releases/latest/
    #
    #               Mic Drop Results (v3.10) [update available]

    console_style(Style.BRIGHT)
    status_msg = f' [{status.value}]' if status else ''
    print(f'Mic Drop Results (v{version_tag}){status_msg}')
    console_style()

    if status != ProgramStatus.UPDATE_AVAILABLE:
    # When an update is available, the download link is already shown above.
    # To avoid confusion, we only print one link at a time.
        print(REPO_URL)


# Section F: Read and process data.xlsx
    xls = pd.ExcelFile(abs_path('data.xlsx'))
    workbook: dict[int | str, pd.DataFrame] = pd.read_excel(
        xls, sheet_name=None)
    xls.close()

    sheet_names = [forbidden_char_pattern.sub('',  # Forbidden file name chars
        str(name)).strip() for name in workbook]
    workbook = {name: list(workbook.values())[i]
                for i, name in enumerate(sheet_names)}


    db_prefix = '__'  # signifies database tables
    database: dict[str, pd.DataFrame] = {}
    for sheet in sheet_names:
        if not sheet.startswith(db_prefix):
            continue  # exclude non-database sheets

        table = workbook[sheet]
        if table.empty or table.shape < (1, 2):  # (1 row, 2 cols) min
            continue
        database[sheet] = table


    groups: dict[str, pd.DataFrame] = {}
    for sheet in sheet_names:
        if sheet.startswith(db_prefix):
            continue  # exclude database tables

        df = workbook[sheet]
        if df.empty or df.shape < (1, n_scols):  # (1 row, n sorting cols) min
            continue

        scols = df.columns.tolist()[:n_scols]  # get sorting columns
        SHEET_INFO = (
            f'{bold("Sheet name:")}  {sheet}\n\n'
            f'See the following row(s) in data.xlsx to find out what '
            f'caused the problem:'
        )


        # Exclude sheets with non-numeric sorting columns
        if any(df.loc[:, scol].dtype.kind not in 'biufc' for scol in scols):
            # Get list of non-numeric values
            str_vals = (df.loc[:, scols][~df.loc[:, scols].applymap(np.isreal)]
                .melt(value_name='__value').dropna()['__value'].tolist())

            Error(60).throw(
                SHEET_INFO,

                preview_df(
                    df, ~df.loc[:, scols].applymap(np.isreal).all(1),
                    n_cols=n_scols, words_to_highlight=str_vals),

                err_type=ErrorType.ERROR)

        # Fill nan values within the sorting columns
        if df.loc[:, scols].isnull().values.any():
            Error(61).throw(
                SHEET_INFO,

                preview_df(
                    df, df.loc[:, scols].isnull().any(axis=1),
                    n_cols=n_scols, words_to_highlight=[None]),

                err_type=ErrorType.WARNING)

            df.loc[:, scols] = df.loc[:, scols].fillna(0)

        # Rank data
        df['__r'] = (
            pd.DataFrame(df.loc[:, scols]            # select sorting columns
            * (np.array(cfg.sort_orders)*2 - 1))  # map bool 0/1 to -1/1
            .apply(tuple, axis=1)  # type: ignore
            .rank(method='min', ascending=False)
            .astype(int))

        # Sort the slides
        df = df.sort_values(by='__r', ascending=True)

        # Remove .0 from whole numbers
        format_int = lambda x: str(int(x)) if x % 1 == 0 else str(x)
        df.loc[:, df.dtypes == float] = (
            df.loc[:, df.dtypes == float].applymap(format_int))

        # Replace {__sheet} with sheet name
        df['__sheet'] = sheet


        # Merge contestant database
        if database:
            process_str = lambda series: (
                series.apply(clean_name) if(series.dtype.kind == 'O')
                else series)

            for table in database.values():
                df_cols = df.columns.tolist()
                db_cols = table.columns.tolist()

                anchor_col = db_cols[0]
                overlapped_cols = [col for col in db_cols
                                   if (col in df_cols) and (col != anchor_col)]

                if anchor_col not in df_cols:  # TODO: add a warning
                    continue

                # Copy processed values of anchor column to '__merge_anchor'
                df['__merge_anchor'] = process_str(df[anchor_col])
                table['__merge_anchor'] = process_str(table[anchor_col])
                table = table.drop(columns=anchor_col)
                table = table.drop_duplicates('__merge_anchor')

                # Merge
                df = df.merge(table, on='__merge_anchor', how='left')
                # Note: merging wtih an existing column will produce duplicates

                for col in overlapped_cols:
                    df[col] = df[f'{col}_y'].fillna(df[f'{col}_x'])
                    df = df.drop(columns=[f'{col}_x', f'{col}_y'])

                df = df.drop(columns='__merge_anchor')


        if '__uid' not in df.columns:
            avatar_mode = False

        df['__template'] = df['__template'].fillna(1)
        df['__uid'] = df['__uid'].str.replace('_', '').str.strip()
        groups[sheet] = df

    if not groups:
        Error(68).throw()


# Section G: Generate PowerPoint slides
    run('TASKKILL /F /IM powerpnt.exe',  # kill all PowerPoint instances
        stdout=DEVNULL, stderr=DEVNULL)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(AVATAR_DIR, exist_ok=True)
    os.makedirs(TEMP_DIR, exist_ok=True)
    check_call(['attrib', '+H', TEMP_DIR])  # hide temp folder

    # Create Module1.bas
    with open(abs_path(TEMP_DIR, 'Module1.bas'), 'w') as f:
        f.write(module1_bas)


    if avatar_mode:
        last_clear_path = abs_path(TEMP_DIR, 'last_clear_avatar_cache.txt')

        try:
            with open(last_clear_path, 'r') as f:
                last_clear_time = int(f.readline())
        except (FileNotFoundError, ValueError):
            last_clear_time = 0

        if time.time() - last_clear_time > 1800*12:  # clear every 12 hours
            for avatar_path in os.scandir(AVATAR_DIR):
                os.unlink(avatar_path)

            with open(last_clear_path, 'w') as f:
                f.write(str(int(time.time())))  # update last clear time

        _import_avatars()


    print('\n\nGenerating slides...')
    print('Please do not click on any PowerPoint window that may '
          'appear during the process.\n')

    for sheet, df in groups.items():
        bar = ProgressBar(
            8, title=sheet, max_title_length=max(map(len, groups.keys())))


        bar.set_description('Opening template.pptm')
        ppt = win32com.client.Dispatch('PowerPoint.Application')
        ppt.Presentations.Open(abs_path('template.pptm'))
        bar.add()


        bar.set_description('Importing macros')
        try:
            ppt.VBE.ActiveVBProject.VBComponents.Import(
                abs_path(TEMP_DIR, 'Module1.bas'))
        except com_error as e:  # trust access not yet enabled
            if e.hresult == -2147352567:  # type: ignore
                Error(41).throw()
            else:
                raise e
        bar.add()


        bar.set_description('Duplicating slides')
        slides_count = ppt.Run('Count')

        # Check for invalid template IDs
        if unknown_templates := [
            x for x in df['__template']
            if as_type(int, x) not in range(1, slides_count + 1)
        ]:
            showcase_cols = ['__r', '__template']
            df_showcase = df[
                showcase_cols
                + [col for col in df.columns if col not in showcase_cols]]
            df_showcase = (df_showcase.drop_duplicates('__template')
                                      .reset_index(drop=True))

            Error(71).throw(
                preview_df(
                    df_showcase, n_cols=2, n_cols_ext=0,
                    words_to_highlight=unknown_templates))

        # Duplicate template slides
        for template in df['__template']:
            ppt.Run('Duplicate', template)
        bar.add()
        ppt.Run(  # delete initial template slides when done
            'DelSlide', *range(1, slides_count + 1))
        bar.add()


        bar.set_description('Saving templates')
        output_path = abs_path(OUTPUT_DIR, f'{sheet}.pptx')
        ppt.Run('SaveAs', str(output_path))
        ppt.Quit()
        bar.add()


        bar.set_description('Filling in judging data')
        prs = Presentation(str(output_path))
        bar.add()
        for i, slide in enumerate(prs.slides):
            fill_slide(slide, {
                k.lstrip('__') : str(v)  # program-generated vars start with __
                for k, v in df.iloc[i].to_dict().items()
            })
        bar.add()


        bar.set_description(f'Saving as {output_path}')
        prs.save(output_path)
        bar.add()


# Section H: Launch the file
    print(f'\nExported to {OUTPUT_DIR}')

    enable_console()

    inp('Press Enter to open the output folder...')
    os.startfile(OUTPUT_DIR)
