from collections.abc import Callable
import configparser
import re
from typing import Any, TypeVar

from exceptions import Error
from utils import abs_path


class ConfigVarTypes:
    update_check: bool
    ini_version_tag: str

    avatar_mode: bool
    last_clear_avatar_cache: int

    trigger_word: str
    ranges: list[float]
    scheme: list[str]
    scheme_alt: list[str]


class Config(ConfigVarTypes):
    def _validate(self):
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


    def __init__(self, filepath: str):
        parser = configparser.ConfigParser()
        parser.read(filepath)

        # Flatten config dict
        self.config: dict[str, Any] = {
            k: v for d in parser.values() for k, v in d.items()
        }

        self._check_missing_vars()   # Check for missing config variables
        self._parse_config()         # Parse values to their assigned types
        self.__dict__ = self.config  # Assign config to class attributes

        try:
            self._validate()         # Validate special conditions
        except AssertionError as e:
            Error(31.1).throw(*e.args)


    def _check_missing_vars(self):
        if missing_vars := [
            v for v in ConfigVarTypes.__annotations__ if v not in self.config
        ]:
            Error(30).throw(str(missing_vars))


    def _parse_config(self):
        for var, var_type in ConfigVarTypes.__annotations__.items():
            try:
                if var_type in [float, str]:
                    self.config[var] = var_type(self.config[var])
                
                elif var_type in [int, bool]:
                    # Fix str conversion issues such as '0' == True
                    self.config[var] = var_type(float(self.config[var]))

                else:  # Remaining is <class 'list'>
                    self.config[var] = self._parse_list(
                        var_type.__args__[0],  # Extract the type of the list
                                               # ... e.g. <class 'float'> if
                                               # ... var_type is list[float]
                        self.config[var])
            except ValueError:
                if var_type.__name__ == 'list':
                    type_name = f'list of {var_type.__args__[0].__name__}'
                    # e.g. 'list of float'
                else:
                    type_name = var_type.__name__
                    # e.g. 'float'

                Error(31).throw(
                    f'Failed to convert the following '
                    f'variable into type: <{type_name}>'
                    f'{self._show_var(var)}'
                )

    T = TypeVar('T')


    def _parse_list(self, var_type: Callable[[str], T], val: str) -> list[T]:
        raw_list = (
            val
            .replace('(', '')
            .replace(')', '')
            .split(','))

        return [var_type(v.strip()) for v in raw_list]


    def _show_var(self, *vars: str) -> str:
        l = [f'    {var} = {self.config[var]}' for var in vars]
        return '\n\n' + '\n'.join(l) + '\n'


if __name__ == '__main__':
    config = Config(abs_path('settings.ini')).__dict__
    print(config)
