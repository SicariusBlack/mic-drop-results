from enum import Enum, auto
import sys
from traceback import format_exception

from colorama import Fore, Style

from constants import *
from utils import inp, console_style


class Traceback:
    SCREENSHOT = ('Please take a screenshot of everything displayed below '
                  'when filling out a bug report. Thank you for your '
                  'patience in getting the issue resolved.')
    
    CONFIG_FORMAT = ('Please verify that the config variables are in their '
                     'valid format according to the note above each variable.')

    _err_lookup = {
    # 0 – 19: Dev-only errors
        0: [
            'Unhandled error.', SCREENSHOT
        ],
        1: [
            'Traceback ID lookup error.',
            'Failed to fetch info from the following traceback ID.'
        ],

    # 20 – 29: API errors
        20: [
            'Failed to communicate with Discord API.',
            'We are unable to download profile pictures at the moment. '
            'Please check your internet connection and try again.'
        ],
        21: [
            'No valid API token found.',
            'Please add a bot token to token.txt or turn off '
            'avatar_mode in settings.ini.'
        ],
        21.1: [
            'Unable to fetch data from Discord API using the following token.',
            'Please replace the following bot token in token.txt with a new '
            'valid one or disable avatar_mode in settings.ini.'
        ],
        22: [
            'Unknown API error.', SCREENSHOT
        ],
        23: [
            'Failed to download profile pictures of the following IDs.',
            'Please check if these user IDs are valid.'
        ],
    # 30 - 39: Config errors
        30: [
            'Missing variable in settings.ini.',
            'The following config variables are missing. Please download '
            'the latest version of settings.ini and try again.\n'
            + LATEST_RELEASE_URL
        ],
        31: [
            'Invalid data type for the following variable in settings.ini.',
            CONFIG_FORMAT
        ],
        31.1: [
            'A variable from settings.ini failed the requirement check.',
            CONFIG_FORMAT
        ],

    # 40 and above: Program errors
        40: [
            'The following files are missing.',
            'Please review the documentation for more information '
            'regarding file requirements.'
        ],
        41: [
            'Failed to import VBA macro due to trust access settings.',
            'Please open PowerPoint, navigate to:\n'
            'File > Options > Trust Center > Trust Center Settings '
            '> Macro Settings, and make sure "Trust access to the VBA '
            'project object model" is enabled.'
        ],
    }

    def lookup(self, tb: float) -> list[str]:
        try:
            return self._err_lookup[tb]
        except KeyError:
            tb_list = [i for i in self._err_lookup if abs(tb - i) < 2]
            Error(1).throw(
                str(tb),
                f'Perhaps you are looking for: {tb_list}',
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
            console_style(Style.BRIGHT)  # Make the error type stand out

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
