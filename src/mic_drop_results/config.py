from collections.abc import Callable
import configparser
import re
from typing import Any, TypeVar

from exceptions import Error


class ConfigVarTypes:
    update_check: bool
    ini_version_tag: str

    sorting_columns: list[bool]

    avatar_mode: bool
    last_clear_avatar_cache: int

    trigger_word: str
    ranges: list[float]
    scheme: list[str]
    scheme_alt: list[str]


class Config(ConfigVarTypes):  # TODO: Add docstrings
    def __init__(self, filepath: str):
        parser = configparser.ConfigParser()
        parser.read(filepath)

        # Flatten config dict
        self.config: dict[str, Any] = {
            k: v for d in parser.values() for k, v in d.items()
        }

        self._check_missing_vars()   # Check for missing config variables
        self._parse_config()         # Get values to their assigned types
        self.__dict__ = self.config  # Assign config to class attributes

        try:
            self._validate()         # Validate special conditions
        except AssertionError as e:
            Error(31.1).throw(*e.args)

    def _validate(self) -> None:
        assert len(self.trigger_word) > 0, (
            'Config variable trigger_word must not be empty.')

        assert (len(self.ranges) ==
                len(self.scheme) ==
                len(self.scheme_alt)), (
            'Lists from variables ranges, scheme, and scheme_alt must '
            'all have the same length.')

        valid_hex = r'^(?:[0-9a-fA-F]{3}){1,2}$'

        assert all(re.fullmatch(valid_hex, h)
                   for scheme in [self.scheme, self.scheme_alt]
                   for h in scheme), (
            'Invalid hex codes found in:'
            + self._show_var('scheme', 'scheme_alt'))

    def _check_missing_vars(self) -> None:
        if missing_vars := [
            v for v in ConfigVarTypes.__annotations__ if v not in self.config
        ]:
            Error(30).throw(str(missing_vars))

    def _parse_config(self) -> None:  # sourcery skip: list-literal
        for name, var_type in ConfigVarTypes.__annotations__.items():
            try:
                match var_type():
                    case str() | float():
                        self.config[name] = var_type(self.config[name])

                    case int() | bool():
                    # Fix conversion issues such as '0' == True
                        self.config[name] = var_type(float(self.config[name]))

                    case list():
                        self.config[name] = self._parse_list(
                            var_type,
                            self.config[name])

            except ValueError:
                if var_type() == list():
                    type_name = f'list of {var_type.__args__[0].__name__}'
                else:
                    type_name = var_type.__name__

                Error(31).throw(
                    f'Failed to convert the following '
                    f'variable into type: <{type_name}>'
                    f'{self._show_var(name)}'
                )

    def _parse_list(self, list_type: Callable[[str], Any], val: str) -> list:
        ele_type = list_type.__args__[0]  # Extract the elements' type
                                          # ... e.g. <class 'float'> if
                                          # ... list_type is list[float]
        raw_list = (val
                    .replace('(', '')
                    .replace(')', '')
                    .split(','))

        match ele_type():
            case int():
                ele_type = lambda v: int(float(v))
            case bool():
                ele_type = lambda v: bool(float(v))

        return [ele_type(element.strip()) for element in raw_list]

    def _show_var(self, *vars: str) -> str:
        l = [f'    {name} = {self.config[name]}' for name in vars]
        return '\n\n' + '\n'.join(l) + '\n'
