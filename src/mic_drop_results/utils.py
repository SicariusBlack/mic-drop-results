from collections.abc import Callable, Generator
from ctypes import windll
import os
import sys
from typing import Any, TypeVar

from colorama import Style
import cursor

from constants import *


# Section A: Basic operations
def is_number(val: Any) -> bool:
    """Checks if value is a number."""
    try:
        float(val)
        return True
    except ValueError:
        return False


T = TypeVar('T')  # Pronounces "typed"
A = TypeVar('A')  # Pronounces "anything"


def as_type(
        t: Callable[[A], T],
        val: A) -> T | A:
    """Returns value as type t if possible, otherwise returns value as it is.

    Args:
        t: the resulting type class to convert.
        val: the value to convert.

    Examples:
        >>> as_type(float, '004'), as_type(int, 3.2)
        (4.0, 3)

        Returns the value as is if it cannot be converted.

        >>> as_type(int, 'banz')
        'banz'
    
    Note:
        Please note that a float under str type (e.g. '3.0') cannot be
        convert directly into type int.

        You could use a wrapper function for t in such cases:

        >>> to_int = lambda str_val: int(float(str_val))
        >>> as_type(to_int, '3.0')
        3
    """
    try:
        return t(val)
    except ValueError:
        return val


def hex_to_rgb(hex_val: str) -> tuple[int, int, int]:
    """Returns in a tuple the RGB values of a color from a given hex code."""
    return tuple(int(hex_val.lstrip('#')[i : i+2], 16) for i in (0, 2, 4))


def parse_version(*versions: str) -> Generator[tuple[int, ...], None, None]:
    """Parses version string into tuple (e.g. 'v3.11.1' into (3, 11, 1)).

    Examples:
        >>> parse_version('3.11.1')
        (3, 11, 1)

        You could also parse multiple versions at the same time:

        >>> latest, current = parse_version('v3.11', 'v3.10')
        >>> current
        (3, 10)
    """
    return (tuple(map(int, v.lstrip('v').split('.'))) for v in versions)


# Section B: File path operations
def abs_path(*rels: str) -> str:
    """Returns absolute path from a relative path.

    Relative path here uses the path to the running file as a reference
    point instead of the current working directory.

    Examples:
        >>> abs_path('vba', 'macros.py')
        'D:\\parent_dir\\src\\md_results\\vba\\macros.py'
    """
    return os.path.join(APP_DIR, *rels)


# Section C: Console utils
def inp(*args: str, **kwargs) -> str:  # TODO: Add docstring, optimize code
    """A wrapper function of the built-in input function.

    This function inherits all the arguments and keyword arguments of
    the built-in print function. Besides, it also enables QuickEdit to
    allow the user to copy printed messages, which are usually error
    details, and disables it thereafter.

    Returns:
        The user input as string.
    """
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


def console_style(*style: str) -> None:
    """Sets the color and style in which the next line is printed.
    
    Args:
        style (optional): an ANSI sequence from the Fore, Back, or Style
            class of the colorama package.

        Pass no argument to reset all formatting.

    Examples:
        Please note that formatting will stack instead of starting anew
        every time you call the function, which means:

        >>> console_style(Fore.RED)
        >>> console_style(Style.BRIGHT)

        ...is equivalent to:

        >>> console_style(Fore.RED, Style.BRIGHT)

        To reset the formatting to default:

        >>> console_style()
    """
    if not style:
        print(Style.RESET_ALL, end='')
    else:
        print(*style, sep='', end='')


class ProgressBar:
    """Creates and prints a progress bar.

    Attributes:
        progr: number of work done. Updates via the add() method.
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
        self.progr: int = 0
        self.total = total
        self.title = title
        self.max_title_length = max_title_length
        self.bar_length = bar_length
        self.desc: str = ''

    def refresh(self) -> None:
        """Reprints the progress bar with updated parameters."""
        filled_length = round(self.bar_length * self.progr / self.total)

        percents = round(100 * self.progr / self.total, 1)
        bar = '█' * filled_length + ' ' * (self.bar_length - filled_length)

        if self.progr > 0:
            sys.stdout.write('\033[2K\033[A\r')  # Delete line, move cursor up,
                                                 # ... and to beginning of line
            sys.stdout.flush()

        title_right_padding = self.max_title_length - len(self.title) + 1
        sys.stdout.write(f'{self.title}{" " * title_right_padding}'
                         f'|{bar}| {self.progr}/{self.total} [{percents}%]'
                         f'{self.desc}')


        # Preview:      Merge   |████████████████████████| 7/7 [100%]
        #               Group 1 |███████████████         | 5/8 [63%]
        #               Filling in judging data


        if self.progr == self.total:
            sys.stdout.write('\033[2K\r')        # Delete line and move cursor
                                                 # ... to beginning of line

        sys.stdout.flush()
        
    def set_description(self, desc: str = '') -> None:
        """Sets the description shown below the progress bar."""
        self.desc = '\n' + desc
        self.refresh()

    def add(self, increment: int = 1) -> None:
        """Updates the progress by a specified increment."""
        self.progr += increment
        self.progr = min(self.progr, self.total)
        self.refresh()
