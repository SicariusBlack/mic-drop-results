from ctypes import windll
import sys
from typing import Any, Generator

from colorama import Style
import cursor


# Section A: Basic operations
def is_number(a: Any) -> bool:
    """Checks if value is a number."""
    try:
        float(a)
        return True
    except ValueError:
        return False


def as_int(a: Any) -> int | Any:
    """Returns value as integer if possible, otherwise returns value as is.

    Examples:
        >>> as_int('banz04')
        'banz04'
        >>> as_int('007')
        7
    """
    try:
        return int(a)
    except ValueError:
        return a


def hex_to_rgb(hex_val: str) -> tuple[int, int, int]:
    """Returns in a tuple the RGB values of a color from a given hex code."""
    return tuple(int(hex_val.lstrip('#')[i : i+2], 16) for i in (0, 2, 4))


def parse_version(*versions: str) -> Generator[tuple[int, ...], None, None]:
    """Parse versions into tuples (e.g. 'v3.11.1' into (3, 11, 1)).

    Examples:
        >>> ver = parse_version('v3.11.1')
        >>> ver
        (3, 11, 1)

        You can also parse multiple versions at the same time:

        >>> latest, current = parse_version('v3.11.1', 'v3.10')
    """
    return (tuple(map(int, v.lstrip('v').split('.'))) for v in versions)


# Section B: Console utils
def inp(*args, **kwargs):  # TODO: Add docstring and optimize code
    # Enable QuickEdit, thus allowing the user to copy printed messages
    kernel32 = windll.kernel32
    kernel32.SetConsoleMode(
        kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x40|0x100))
    cursor.show()

    print(*args, **kwargs, end='')
    i = input()

    # Disable QuickEdit
    kernel32.SetConsoleMode(
        kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x00|0x100))
    cursor.hide()

    return i


def console_style(style: str = Style.RESET_ALL) -> None:
    """Sets the color and style in which the next line is printed.
    
    Args:
        color (optional): an ANSI sequence from the Fore, Back, or Style
            class of the colorama package.

        Pass no argument to reset all formatting.

    Examples:
        >>> console_style(Fore.RED)
        >>> console_style(Style.BRIGHT)

        To reset the style to default:

        >>> console_style()
    """
    print(style, end='')


class ProgressBar:
    """Creates and prints a progress bar.

    Attributes:
        progress: number of work done. Updates via the add() method.
        total: number of work to perform.
        title: title shown to the left of the progress bar.
        max_title_length: length of the longest title to ensure left
            alignment of the progress bars when there are more than
            one bar.
        bar_length: length of the progress bar in characters.
        desc: description of the task in progress shown below the
            progress bar. Updates via the set_description() method.
    """

    def __init__(self, total: int, title: str, max_title_length: int,
                 bar_length: int = 40) -> None:
        self.progress: int = 0
        self.total = total
        self.title = title
        self.max_title_length = max_title_length
        self.bar_length = bar_length
        self.desc: str = ''

    def refresh(self) -> None:
        """Reprints the progress bar with updated parameters."""
        filled_length = round(self.bar_length * self.progress / self.total)

        percents = round(100 * self.progress / self.total, 1)
        bar = '█' * filled_length + ' ' * (self.bar_length - filled_length)

        if self.progress > 0:
            sys.stdout.write('\033[2K\033[A\r')  # Delete line, move cursor up,
                                                 # and to beginning of the line
            sys.stdout.flush()

        title_right_padding = self.max_title_length - len(self.title) + 1
        sys.stdout.write(f'{self.title}{" " * title_right_padding}'
                         f'|{bar}| {self.progress}/{self.total} [{percents}%]'
                         f'{self.desc}')


        # Preview:      Merge   |████████████████████████| 7/7 [100%]
        #               Group 1 |███████████████         | 5/8 [63%]
        #               Filling in judging data


        if self.progress == self.total:
            sys.stdout.write('\033[2K\r')        # Delete line and move cursor
                                                 # to beginning of line

        sys.stdout.flush()
        
    def set_description(self, desc: str = '') -> None:
        """Sets the description shown below the progress bar."""
        self.desc = '\n' + desc
        self.refresh()

    def add(self, increment: int = 1) -> None:
        """Updates the progress by a specified increment."""
        self.progress += increment
        self.progress = min(self.progress, self.total)
        self.refresh()
