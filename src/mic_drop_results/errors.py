import copy
from enum import Enum, auto
import sys
from traceback import format_exception

from colorama import Fore, Style

from constants import *
from utils import abs_path, inp, console_style


class Tag(Enum):
    DEV = 'DEV'
    SYS = 'SYSTEM'
    INTERNET = 'ConnectionError'
    SETTINGS_INI = 'settings.ini'
    TOKEN_TXT = 'token.txt'
    DATA_XLSX = 'data.xlsx'


class Traceback:
    templates = {
        'screenshot': [
            'Please take a screenshot of everything displayed below when '
            'you fill out a bug report. Thank you for your patience in '
            'getting this issue resolved.'],

        'cfg_format': [
            'Please verify that the config variables are in their '
            'valid format according to the note above each variable.'],
    }

    _err_lookup = {
    # 0 – 19: Dev errors
        0: [
            Tag.DEV, 'Unhandled error.'
        ],
        1: [
            Tag.DEV, 'Traceback ID lookup error.'
        ],

    # 20 – 29: API errors
        20: [
            Tag.INTERNET, 'Failed to communicate with Discord API.',
            'We are unable to download avatars at the moment. '
            'Please check your internet connection and try again.'
        ],
        21: [
            Tag.TOKEN_TXT, 'No valid bot token found.',
            'Please add a bot token to token.txt or turn off '
            'avatar_mode completely in settings.ini.'
        ],
        21.1: [
            Tag.TOKEN_TXT, 'Invalid bot token.',
            'The following bot token is invalid or has expired. '
            'Please replace the following token from token.txt with a '
            'new valid one.'
        ],
        22: [
            Tag.DEV, 'Unknown Discord API error.'
        ],
        23: [
            Tag.DATA_XLSX,
            'Unable to download the avatars of certain users.',
            'Please make sure the following user IDs are correctly typed.'
        ],

    # 30 – 39: Config errors
        30: [
            Tag.SETTINGS_INI, 'Missing config variables.',
            'The following config variables are missing. Please download '
            'the latest version of settings.ini and try again.\n'
            + TEMPLATES_URL
        ],
        31: [
            Tag.SETTINGS_INI, 'Invalid data type for config variable.',
            *templates['cfg_format']
        ],
        31.1: [
            Tag.SETTINGS_INI, 'Config variable failed requirement check.',
            *templates['cfg_format']
        ],

    # 40 – 59: System errors
        40: [
            Tag.SYS, 'The following files are missing.',
            'Please download the missing files from the following source.\n'
            + TEMPLATES_URL
        ],
        41: [
            Tag.SYS, 'Failed to import VBA macros due to privacy settings.',
            'Please open PowerPoint and navigate to:\nFile > Options '
            '> Trust Center > Trust Center Settings > Macro Settings',
            'Make sure the "Trust access to the VBA project object model" '
            'option is checked.'
        ],

    # 60 and above: Data errors
        60: [
            Tag.DATA_XLSX, 'Sorting columns contain text.',
            'The sorting columns of the following sheet contain text '
            'but expects numeric data type.\n'
            'Have you pasted data in the wrong column, by any chance?'
        ],
        61: [
            Tag.DATA_XLSX, 'Sorting columns contain empty value.',
            'The sorting columns of the following sheet contain empty '
            'cell values.\n'
            'These empty values will be replaced with 0 if you proceed on.'
        ],
        68: [
            Tag.DATA_XLSX, 'No valid sheet found.',
            'We have examined every sheet from the following Excel file:\n'
            + abs_path("data.xlsx"),
            'No sheet appears to be in the correct format.',
            'Please download a sample data.xlsx file from the following '
            'URL and use it as a reference for customizing your own.\n'
            + TEMPLATES_URL
        ],
        70: [
            Tag.DATA_XLSX, 'No leading underscore before user IDs.',
            'Please add an underscore (_) before every user ID from the '
            '"__uid" column. For example: _1010885414850154587',
            'This is to prevent Excel and the program from rounding the UIDs.',
            'Make sure all IDs are still valid before adding the underscores.'
        ],
        71: [
            Tag.DATA_XLSX, 'Template does not exist.',
            'The following template(s) could not be matched with any '
            'slide from template.pptm.'
        ],
    }

    def lookup(self, tb: float) -> list[str]:
        try:
            res = copy.deepcopy(self._err_lookup[tb])
            res[1] = f'[{res[0].value}] {res[1]}'

            if res[0] == Tag.DEV:
                res += self.templates['screenshot']

            return res[1:]
        except KeyError:
            tb_list = [i for i in self._err_lookup if abs(tb - i) < 5]
            if not tb_list:
                tb_list = list(self._err_lookup)

            Error(1).throw(
                f'Traceback ID: {tb}\n'
                f'Perhaps you are looking for: {tb_list}'
            )
            return []


class ErrorType(Enum):
    ERROR = auto()
    WARNING = auto()
    INFO = auto()


class Error(Traceback):
    def __init__(self, tb: float):
        self.tb = tb
        self.tb_code = self.get_code()

        # Look up content with traceback ID
        self.content: list[str] = super().lookup(tb)

    def get_code(self) -> str:
        whole, decimal = int(self.tb), round(self.tb%1, 7)

        whole = str(whole).zfill(3)
        decimal = str(decimal).replace('.', '')

        code = (whole
                if int(decimal) == 0
                else f'{whole}.{int(decimal)}')
        return f'E-{code}'

    def throw(
            self, *details: str, err_type: ErrorType = ErrorType.ERROR
        ) -> None:
        self.content += [*details]
        self._print(*self.content, err_type=err_type)

    def _print(
            self, *content: str,  err_type: ErrorType = ErrorType.ERROR
        ) -> None:
        """Handles and reprints an error with human-readable details.

        Prints an error message with paragraphs explaining the error
        and double-spaced between paragraphs.

        The first paragraph will be shown beside the error type and will
        inherit the color red if it is an error, the color yellow if it
        is a warning, and the default color if it is an info message.

        Args:
            *content: every argument makes a paragraph of the error
                message. The first paragraph should summarize the error
                in one sentence. The rest of the paragraphs will explain
                what causes and how to resolve the error.
            err_type (optional): the error type taken from the ErrorType
                class. Defaults to ErrorType.ERROR.
        """
        if content:
            console_style(Style.BRIGHT)

            if err_type == ErrorType.ERROR:
                console_style(Fore.RED)
            elif err_type == ErrorType.WARNING:
                console_style(Fore.YELLOW)

            print(f'\n\n{err_type.name}:{Style.NORMAL} {content[0]} '
                  f'(Traceback code: {self.tb_code})')
            console_style()

        if len(content) > 1:
            print()
            print(*content[1:], sep='\n\n')

        if err_type == ErrorType.ERROR:
            inp('\nPress Enter to exit the program...')
            sys.exit(1)
        else:
            inp('\nPress Enter to continue...')


def print_exception_hook(exc_type, exc_value, tb) -> None:
    Error(0).throw(''.join(format_exception(exc_type, exc_value, tb))[:-1])
