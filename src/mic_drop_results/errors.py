# Copyright 2023 Phan Huy

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.

import copy
from enum import Enum, auto
import os
from traceback import format_exception

from rich.padding import Padding

from compiled_regex import *
from constants import *
from utils import inp, abs_dir


class Tag(Enum):
    DEV = "DEV"
    SYS = "SYSTEM"
    INTERNET = "ConnectionError"
    FILE_SETTINGS = "settings.ini"
    FILE_TOKEN = "token.txt"
    FILE_DATA = "data.xlsm"


class Traceback:
    templates = {
        "screenshot": [
            "To help us fix this bug, please capture and attach a screenshot of everything you see below. "
            "We appreciate your cooperation in helping us resolve this issue.",
        ],
        "invalid_config": [
            "Make sure that these configs follow the valid formats as explained in the notes from settings.ini:"
        ],
        "developer_portal": [
            "Visit Discord's Developer Portal and follow the steps in token.txt to create your token.\n"
            "https://discord.com/developers/applications"
        ],
    }

    _err_lookup = {
        # 0 – 19: Dev errors
        0: [Tag.DEV, "Unhandled error."],
        1: [Tag.DEV, "Traceback ID lookup error."],
        # 20 – 29: API errors
        20: [
            Tag.INTERNET,
            "Failed to communicate with Discord's API",
            "We are unable to download avatars at the moment. Please check your internet connection and try again.",
        ],
        21: [
            Tag.FILE_TOKEN,
            "No valid bot token found",
            "You need to add your token(s) in token.txt or disable avatar mode in settings.ini.",
            *templates["developer_portal"],
        ],
        21.1: [
            Tag.FILE_TOKEN,
            "Invalid bot token",
            "Your bot token is either invalid or has been deactivated. Please replace this token from token.txt with a new valid one:",
            *templates["developer_portal"],
        ],
        22: [Tag.DEV, "Unknown Discord's API error."],
        23: [
            Tag.FILE_DATA,
            "Failed to fetch the data of certain users",
            "We are unable to download the avatars of these users:",
            "Verify that these IDs are valid and that the corresponding accounts have not been deleted.",
        ],
        # 30 – 39: Config errors
        30: [
            Tag.FILE_SETTINGS,
            "Missing config",
            "You are using an outdated settings file that does not have the required config:",
            "To fix this, download and use the latest settings.ini file from this source:\n"
            + TEMPLATES_URL,
        ],
        31: [
            Tag.FILE_SETTINGS,
            "Invalid data type for config",
            *templates["invalid_config"],
        ],
        31.1: [
            Tag.FILE_SETTINGS,
            "Config failed requirement check",
            *templates["invalid_config"],
        ],
        # 40 – 59: System errors
        40: [
            Tag.SYS,
            "Missing required files",
            "This program cannot run without these essential files:",
            "Please download the missing files from this source and paste them into your working directory.\n"
            + TEMPLATES_URL,
        ],
        41: [
            Tag.SYS,
            "Failed to import VBA macros due to a privacy settings",
            "Please open PowerPoint and navigate to:\n"
            "File > Options > Trust Center > Trust Center Settings > Macro Settings\n\n"
            'Make sure the "Trust access to the VBA project object model" option is checked.',
        ],
        # 60 and above: Data errors
        60: [
            Tag.FILE_DATA,
            "Sorting columns cannot contain text",
            "The following sheet has text values in the columns that should only contain numbers. This may cause errors in the sorting process.\n"
            "Please check if you have entered data in the correct columns.",
        ],
        61: [
            Tag.FILE_DATA,
            "Sorting columns cannot contain empty values",
            "Some cells in the columns that you want to sort are blank. This may affect the accuracy of the sorting process.\n"
            "The system will automatically fill these cells with zeros if you continue. Please confirm if you wish to proceed.",
        ],
        68: [
            Tag.FILE_DATA,
            "No valid sheet found",
            "None of the sheets in data.xlsm have the correct and usable format that we require.",
            "Please download a sample data.xlsm file from this source and use it as a reference to customize your own file.\n"
            + TEMPLATES_URL,
        ],
        70: [
            Tag.FILE_DATA,
            "Missing an underscore before every user ID",
            'To avoid unwanted rounding of the user IDs by Excel or the program, please prefix an underscore (_) to each user ID in the "__uid" column.',
            "For example, the user ID 1104424999365918841 should be written as _1104424999365918841.",
            "This will ensure that the user IDs are treated as text values and not as numbers.",
        ],
        71: [
            Tag.FILE_DATA,
            "Template does not exist",
            "We could not find any matching slides for the following template(s) in template.pptm:",
        ],
    }

    def lookup(self, tb: float) -> list[str]:
        try:
            res = copy.deepcopy(self._err_lookup[tb])
            res[1] = f"[{res[0].value}] {res[1]}"

            if res[0] == Tag.DEV:
                res += self.templates["screenshot"]

            return res[1:]
        except KeyError:
            tb_list = [i for i in self._err_lookup if abs(tb - i) < 5]
            if not tb_list:
                tb_list = list(self._err_lookup)

            Error(1).throw(
                f"Traceback ID: {tb}\n" f"Perhaps you are looking for: {tb_list}"
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
        whole, decimal = int(self.tb), round(self.tb % 1, 7)

        whole = str(whole).zfill(3)
        decimal = str(decimal).replace(".", "")

        code = whole if int(decimal) == 0 else f"{whole}.{int(decimal)}"
        return f"E-{code}"

    def throw(self, *details: str, err_type: ErrorType = ErrorType.ERROR) -> None:
        if len(self.content) >= 3:
            self.content = self.content[:2] + [*details] + self.content[2:]
        else:
            self.content += [*details]

        # Redact sensitive information
        for i, x in enumerate(self.content):
            self.content[i] = match_windows_username.sub("user", str(x))

        self._print(*self.content, err_type=err_type)

    def _print(self, *content: str, err_type: ErrorType = ErrorType.ERROR) -> None:
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
        assert content, "Please provide details on this error."
        assert (
            len(content) >= 2
        ), "An error name and a short description of the error need to be provided."

        style = None
        if err_type == ErrorType.ERROR:
            style = "red"
        elif err_type == ErrorType.WARNING:
            style = "yellow"

        console.print(
            f"\n[bold]{err_type.name}:[/bold] \\{content[0]}"
            + f" (Traceback code: {self.tb_code})",
            style=style,
        )  # error details

        console.line(1)
        console.print(content[1])  # steps to resolve
        console.line(1)

        for part in content[2:]:
            console.print(Padding(part, (1, 4, 0, 4)))  # extra details

        if err_type == ErrorType.ERROR:
            console.line(2)
            inp("Press Enter to exit the program...\n\n")
            os._exit(1)
        else:
            console.line(2)
            inp("Press Enter to skip this warning...\n\n", hide_text=True)


def print_exception_hook(exc_type, exc_value, tb) -> None:
    Error(0).throw("".join(format_exception(exc_type, exc_value, tb))[:-1])
