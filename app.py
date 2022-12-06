import configparser
import contextlib
from ctypes import windll
from io import BytesIO
import itertools
from json import load, dump
from multiprocessing import Pool, freeze_support
import multiprocessing.popen_spawn_win32 as forking
import numpy as np
import os
from PIL import Image
import re
import requests
from signal import signal, SIGINT, SIG_IGN
from subprocess import run, DEVNULL
import sys
import time
from traceback import print_exception
from urllib.request import Request, urlopen
import webbrowser

import cursor
from colorama import init, Fore

import cv2
import pandas as pd

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.slide import Slide
from pptx.util import Inches

from pywintypes import com_error
import win32com.client


class _Popen(forking.Popen):
    def __init__(self, *args, **kw):
        """Makes multiprocessing compatible with pyinstaller.
        Source: https://github.com/pyinstaller/pyinstaller/wiki/Recipe-Multiprocessing
        """
        if hasattr(sys, 'frozen'):
            os.putenv('_MEIPASS2', sys._MEIPASS)
        try:
            super(_Popen, self).__init__(*args, **kw)
        finally:
            if hasattr(sys, 'frozen'):
                if hasattr(os, 'unsetenv'):
                    os.unsetenv('_MEIPASS2')
                else:
                    os.putenv('_MEIPASS2', '')
forking.Popen = _Popen


class ProgressBar:
    def __init__(self, total, bar_length, title, max_title_length):
        self.progress = 0
        self.total = total
        self.bar_length = bar_length
        self.title = title
        self.max_title_length = max_title_length
        self.desc = ''

    def add(self, incr=1):
        self.progress += incr
        self.refresh()

    def refresh(self):
        filled_length = int(round(self.bar_length * self.progress / float(self.total)))

        percents = round(100 * self.progress / float(self.total), 1)
        bar = '█' * filled_length + ' ' * (self.bar_length - filled_length)

        if self.progress > 0:
            sys.stdout.write('\033[2K\033[A\r')  # Delete line, move cursor up, and to beginning of the line
            sys.stdout.flush()

        sys.stdout.write(f'{self.title}{" " * (self.max_title_length - len(self.title))} '
                         f'|{bar}| {self.progress}/{self.total} [{percents}%]{self.desc}')


        # Preview:   Group 1 |███████████████         | 5/8 [63%]
        #            Filling in judging data


        if self.progress >= self.total:
            sys.stdout.write('\033[2K\r')        # Delete line and move cursor to beginning of line

        sys.stdout.flush()
        
    def set_description(self, text):
        self.desc = '\n' + text
        self.refresh()


def is_number(a):
    try:
        float(a)
        return True
    except ValueError:
        return False


def as_int(a):
    try:
        return int(a)
    except ValueError:
        return a


def set_console_color(color=Fore.RESET):
    print(color, end='')


def throw_error(*messages, err_type: str = 'error'):
    """Handles and throws an error with additional guides and details."""
    if messages:
        if err_type == 'error':
            set_console_color(Fore.RED)     # For errors
        else:
            set_console_color(Fore.YELLOW)  # For warnings

        print(f'\n\n{err_type.upper()}: {messages[0]}')
        set_console_color()

    if len(messages) > 1:
        print()
        print(*messages[1:], sep='\n\n')

    if err_type == 'error':
        _input('\nPress Enter to exit the program...')  # For errors
        sys.exit(1)
    else:
        _input('\nPress Enter to continue...')          # For warnings


def print_exception_and_exit(exc_type, exc_value, tb):
    print_exception(exc_type, exc_value, tb)
    throw_error()


def hex_to_rgb(hexcode):
    return tuple(int(hexcode.lstrip('#')[i : i+2], 16) for i in (0, 2, 4))


def _input(*args, **kwargs):
    # Enable QuickEdit, thus allowing the user to copy the error message
    kernel32 = windll.kernel32
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x40|0x100))
    cursor.show()

    print(*args, **kwargs, end='')
    i = input()

    # Disable QuickEdit and Insert mode
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x00|0x100))
    cursor.hide()

    return i


def replace_text(slide: Slide, df, i, avatar_mode) -> Slide:
    """Replaces and formats text."""
    cols = df.columns.values.tolist() + ['p']

    for shape in slide.shapes:
        if not shape.has_text_frame or '{' not in shape.text:
            continue

        text_frame = shape.text_frame

        for run in itertools.chain.from_iterable([p.runs for p in text_frame.paragraphs]):
            for search_str in set(re.findall(r'(?<={)(.*?)(?=})', run.text)).intersection(cols):
                # Profile picture
                if search_str == 'p':
                    # Test cases
                    if 'uid' not in cols or not avatar_mode:
                        continue

                    if pd.isnull(df['uid'].iloc[i]):
                        run.text = ''
                        continue

                    # Extract effect index and remove {p}
                    effect = as_int(run.text.strip()[3:])
                    run.text = ''

                    uid = str(df['uid'].iloc[i]).strip().replace('_', '')

                    og_path = avapath + '_' + uid + '.png'
                    img_path = avapath + str(effect) + '_' + uid + '.png'

                    if not os.path.isfile(og_path):
                        continue

                    if is_number(effect):
                        img = cv2.imread(og_path)
                        match effect:
                            case 1:
                                img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

                        cv2.imwrite(img_path, img)

                    new_shape = slide.shapes.add_picture(
                        img_path, shape.left, shape.top,
                        shape.width, shape.height
                    )

                    new_shape.auto_shape_type = MSO_SHAPE.OVAL
                    old = shape._element
                    new = new_shape._element
                    old.addnext(new)
                    old.getparent().remove(old)
                    continue

                # Actual text
                repl = str(df[search_str].iloc[i])
                repl = repl if repl != 'nan' else ''  # Replace missing values with blank

                run_text = run.text

                if search_str.startswith(starts):
                    run.text = repl
                else:
                    run.text = run.text.replace('{' + search_str + '}', repl)

                # Replace image links
                pattern = r'\<\<(.*?)\>\>'  # Regex pattern to look for <<image_links.png>>
                img_link = re.findall(pattern, run.text)

                if len(img_link) > 0:
                    try:
                        img = BytesIO(requests.get(img_link[0]).content)
                        pil = Image.open(img)

                        im_width = shape.height / pil.height * pil.width
                        new_shape = slide.shapes.add_picture(
                            img, shape.left + (shape.width - im_width) / 2, shape.top,
                            im_width, shape.height
                        )

                        old = shape._element.addnext(new_shape._element)

                        run.text = re.sub(pattern, '', run.text)
                        text_frame.margin_left = Inches(5.2)
                    except Exception:
                        throw_error('Could not load the following image '
                           f'(Slide {i + 1}, {df["sheet"].iloc[0]}).\n{img_link[0]}',
                            'Please check your internet connection and verify that '
                            'the link leads to an image file. '
                            'It should end with an image extension like .png in most cases.',
                            err_type='warning')

                # Color formatting
                if not search_str.startswith(starts):
                    continue

                # Check RGB
                if run.font.color.type == MSO_COLOR_TYPE.RGB and \
                    run.font.color.rgb not in [RGBColor(0, 0, 0), RGBColor(255, 255, 255)]:
                    continue

                for ind, val in enumerate(range_list):
                    if is_number(repl) and float(repl) >= val:
                        if run_text.endswith('1'):
                            run.font.color.rgb = RGBColor(*color_list_light[ind])
                        else:
                            run.font.color.rgb = RGBColor(*color_list[ind])
                        break
    return slide


def get_avatar(id, api_token):
    header = {
        'Authorization': 'Bot ' + api_token
    }

    if not is_number(id):
        return None

    link = None

    try:
        response = requests.get(f'https://discord.com/api/v9/users/{id}', headers=header)
        link = f'https://cdn.discordapp.com/avatars/{id}/{response.json()["avatar"]}'
    except KeyError:
        if response.json()['message'] == '401: Unauthorized':
            throw_error('Invalid token. Please provide a new token in token.txt or '
                        'turn off avatar_mode in config.cfg.', response.json())
        elif response.json()['message'] == 'You are being rate limited.':
            time.sleep(response.json()['retry_after'])
            get_avatar(id, api_token)
        else:
            throw_error(response.json(), err_type='warning')
    except requests.exceptions.ConnectionError:
        throw_error('Unable to connect with Discord API. Please check your internet '
                    'connection and try again.', err_type='warning')
    return link


def download_avatar(uid, avapath, api_token):
    uid = uid.strip().replace('_', '')
    img_path = avapath + '_' + uid.strip() + '.png'

    # Load image from link
    avatar_url = get_avatar(uid, api_token)

    if not avatar_url:
        return None

    avatar_url += '.png'

    req = urlopen(Request(avatar_url, headers={'User-Agent': 'Mozilla/5.0'}))
    arr = np.asarray(bytearray(req.read()), dtype=np.uint8)
    img = cv2.imdecode(arr, -1)

    cv2.imwrite(img_path, img)


if __name__ == '__main__':
    # Section A: Fix Command Prompt issues
    freeze_support()          # Multiprocessing freeze support
    signal(SIGINT, SIG_IGN)   # Handle KeyboardInterrupt

    # Disable QuickEdit and Insert mode
    kernel32 = windll.kernel32
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x00|0x100))

    sys.excepthook = print_exception_and_exit   # Avoid exiting the program when an error is thrown
    init()                                      # Enable ANSI escape sequences
    cursor.hide()                               # Hide cursor


    # Section B: Check for missing files
    if missing := [f for f in [
            'config.cfg', 'data.xlsx', 'template.pptm', 'Module1.bas', 'token.txt'
        ] if not os.path.isfile(f)]:
        throw_error('The following files are missing. Please review the documentation for more '
            'information related to file requirements.', '\n'.join(missing))


    # Section C: Load config.cfg
    config = load(open('config.json'))

    # Store config variables as local variables
    range_list = config['format']['ranges'][::-1]
    color_list = config['format']['colors'][::-1]
    color_list_light = config['format']['colors_light'][::-1]
    starts = config['format']['starts_with']
    avatar_mode = config['avatars']
    last_clear = config['last_clear_avatar_cache']

    with open('token.txt') as f:
        token_list = f.read().splitlines()
        token_list = [i.strip('"') for i in token_list if len(i) > 62]

    if not token_list and avatar_mode:
        throw_error('Please provide a valid bot token in token.txt or turn off avatar mode in config.json.')

    color_list = list(map(hex_to_rgb, color_list))
    color_list_light = list(map(hex_to_rgb, color_list_light))


    # Section D: Check for updates
    status, url = '', ''

    with contextlib.suppress(requests.exceptions.ConnectionError, KeyError):
        if config['update_check']:
            response = requests.get('https://api.github.com/repos/banz04/mic-drop-results/releases/latest', timeout=3)

            raw_ver = response.json()['tag_name'][1:]
            version, config_ver = [tuple(map(int, v.split('.'))) for v in 
                [raw_ver, config['version']]
            ]

            if version > config_ver:
                set_console_color(Fore.YELLOW)
                print(f'Update {raw_ver}')
                print(response.json()['body'].partition('\n')[0])

                url = 'https://github.com/banz04/mic-drop-results/releases/latest/'
                print(url + '\n')
                webbrowser.open(url, new=2)
                set_console_color()

                status = 'update available'
            elif version < config_ver:
                status = 'beta'
            else:
                status = 'latest'

            status = f' [{status}]'

    print(f'Mic Drop Results (v{config["version"]}){status}')
    set_console_color()

    if 'update available' not in status:
        url = 'https://github.com/banz04/mic-drop-results'
        print(url)


    # Section E: Process the data
    folder_path = os.getcwd() + '\\'
    outpath = folder_path + 'output\\'
    avapath = folder_path + 'avatars\\'

    xls = pd.ExcelFile('data.xlsx')

    sheetnames_raw = xls.sheet_names
    sheetnames = [re.sub(r'[\\\/:"*?<>|]+', '', name) for name in sheetnames_raw]
    data = {}

    db_list = []
    for sheet in sheetnames_raw:
        if sheet.startswith('_'):
            df = pd.read_excel(xls, sheet)

            # Validate shape
            if df.empty or df.shape < (1, 2):
                continue

            db_list.append(df)

    SHARING_VIOLATION = '\033[33mNOTE: Please exit the program before modifying data.xlsx or ' \
        'Microsoft Excel will throw a Sharing Violation error.\033[39m'

    for i, sheet in enumerate(sheetnames_raw):
        df = pd.read_excel(xls, sheet)

        # Validate shape
        if df.empty or df.shape < (1, 2):
            continue

        # Exclude database sheets
        if sheet.startswith('_'):
            continue

        # Exclude sheets with first two columns where data types are not numeric
        if sum(df.iloc[:, i].dtype.kind in 'biufc' for i in range(2)) < 2:
            throw_error(f'Invalid data type. The following rows of {sheet} contain strings '
                'instead of the supposed numeric data type within the first two columns. '
                'The sheet will be excluded if you proceed on.',

                df[~df.iloc[:, :2].applymap(np.isreal).all(1)],

                err_type='warning'
            )

            continue

        # Replace NaN values within the first two columns with 0
        if df.iloc[:, :2].isnull().values.any():
            throw_error(f'The following rows of {sheet} contain empty values '
                'within the first two columns.',

                df[df.iloc[:, :2].isnull().any(axis=1)],

                'You may exit this program and modify your data or proceed on with '
                'these empty values substituted with 0.', SHARING_VIOLATION,

                err_type='warning'
            )

            df.iloc[:, :2] = df.iloc[:, :2].fillna(0)

        # Check for cases where avg and std are the same (hold the same rank)
        df['r'] = pd.DataFrame(zip(df.iloc[:, 0], df.iloc[:, 1] * -1)) \
            .apply(tuple, axis=1).rank(method='min', ascending=False).astype(int)

        # Sort the slides
        df = df.sort_values(by='r', ascending=True)

        # Remove .0 from whole numbers
        format_number = lambda x: str(int(x)) if x % 1 == 0 else str(x)
        df.loc[:, df.dtypes == float] = df.loc[:, df.dtypes == float].applymap(format_number)

        # Replace {sheet} with sheet name
        df['sheet'] = sheet

        # Merge contestant database
        clean_name = lambda x: x.str.lower().str.strip() if(x.dtype.kind == 'O') else x
        if db_list:
            for db in db_list:
                df_cols = df.columns.values.tolist()
                db_cols = db.columns.values.tolist()
                merge_col = db_cols[0]

                # Use merge for non-existing columns
                df = df.merge(db[[merge_col] + [i for i in db_cols if i not in df_cols]],
                    left_on=clean_name(df[merge_col]), right_on=clean_name(db[merge_col]), how='left')

                df.loc[:, merge_col] = df[merge_col + '_x']
                df.drop(['key_0', merge_col + '_x', merge_col + '_y'], axis=1, inplace=True)

                # Use update for existing columns
                df['update_index'] = clean_name(df[merge_col])
                df = df.set_index('update_index')

                db['update_index'] = clean_name(db[merge_col])
                db = db.set_index('update_index')
                db_cols.remove(merge_col)

                update_cols = [i for i in db_cols if i in df_cols]
                df.update(db[update_cols])
                df.reset_index(drop=True, inplace=True)

        if 'uid' not in df.columns.values.tolist(): avatar_mode = 0

        # Fill in missing templates
        df['template'].fillna(1, inplace=True)

        data[sheetnames[i]] = df

    if not data:
        throw_error(f'No valid sheet found in {folder_path}data.xlsx')


    # Section F: Generate PowerPoint slides
    print('\nGenerating slides...')
    print('Please do not click on any PowerPoint windows that may show up in the process.\n')

    # Kill all PowerPoint instances
    run('TASKKILL /F /IM powerpnt.exe', stdout=DEVNULL, stderr=DEVNULL)

    # Open template presentation
    os.makedirs(outpath, exist_ok=True)
    os.makedirs(avapath, exist_ok=True)

    # Clear cache
    if time.time() - last_clear > 1800:  # Clears every hour
        for f in os.scandir(avapath):
            os.unlink(f)

        # Update last clear time
        with open('config.json', 'w') as f:
            config['last_clear_avatar_cache'] = int(time.time())
            dump(config, f, indent=4)

    # Download avatars with parallel processing
    attempt = 0
    pool = Pool(3)

    while avatar_mode:
        uid_list = []

        for df in data.values():
            if df['uid'].dtype.kind in 'biufc':
                throw_error('The \'uid\' column has numeric data type instead of the supposed string data type.',
                    'Please exit the program and add an underscore before each user ID.', SHARING_VIOLATION)

            uid_list += [id for id in df['uid'] if not pd.isnull(id) and not os.path.isfile(avapath + id.strip() + '.png')]

        if len(uid_list) == 0:
            break

        if attempt > 0 and attempt <= 3:
            print(f'Unable to download the profile pictures of the following users. Retrying {attempt}/3', uid_list, sep='\n', end='\n\n')
        elif attempt > 3:
            throw_error('Failed to download the profile pictures of the following users. Please verify that their user IDs are correct.', uid_list,
                err_type='warning')

        pool.starmap(download_avatar, zip(uid_list,
            [avapath] * len(uid_list), itertools.islice(itertools.cycle(token_list), len(uid_list))))

        attempt += 1

    pool.close()
    pool.join()

    for k, df in data.items():
        bar = ProgressBar(8, 40, group=k, group_len=max(map(len, data.keys())))

        # Open template presentation
        bar.set_description('Opening template.pptm')
        ppt = win32com.client.Dispatch('PowerPoint.Application')
        ppt.Presentations.Open(f'{folder_path}template.pptm')
        bar.add()

        # Import macros
        bar.set_description('Importing macros')

        try:
            ppt.VBE.ActiveVBProject.VBComponents.Import(f'{folder_path}Module1.bas')
        except com_error as e:
            if e.hresult == -2147352567:
                # Warns the user about trust access error
                throw_error('Please open PowerPoint, look up Trust Center Settings, '
                            'and make sure Trust access to the VBA project object model is enabled.')
            else:
                raise e

        bar.add()

        # Duplicate slides
        bar.set_description('Duplicating slides')
        slides_count = ppt.Run('Count')

        # Duplicate slides
        for t in df.loc[:, 'template']:
            if as_int(t) not in range(1, slides_count + 1):
                throw_error(f'Template {t} does not exist. ({k})',
                    'Please exit the program and modify the \'template\' column of data.xlsx', SHARING_VIOLATION)


            ppt.Run('Duplicate', t)

        bar.add()

        # Delete template slides when done
        ppt.Run('DelSlide', *range(1, slides_count + 1))
        bar.add()

        # Save as output file
        bar.set_description('Saving templates')
        output_filename = f'{k}.pptx'

        ppt.Run('SaveAs', f'{outpath}{output_filename}')
        bar.add()
        run('TASKKILL /F /IM powerpnt.exe', stdout=DEVNULL, stderr=DEVNULL)
        bar.add()

        # Replace text
        bar.set_description('Filling in judging data')
        prs = Presentation(outpath + output_filename)

        for i, slide in enumerate(prs.slides):
            replace_text(slide, df, i, avatar_mode)
        bar.add()

        # Save
        bar.set_description(f'Saving as {outpath + output_filename}')
        prs.save(outpath + output_filename)
        bar.add()


    # Section G: Launch the file
    print(f'\nExported to {outpath}')

    # Enable QuickEdit
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x40|0x100))

    _input('Press Enter to open the output folder...')
    os.startfile(outpath)
