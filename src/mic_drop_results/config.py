from collections.abc import Callable
import configparser
import re
from typing import TypedDict, Any, TypeVar

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
    def validate(self):
        assert len(self.trigger_word) > 0, (
            'Variable trigger_word must not be empty.')
        
        assert len(self.ranges) == len(self.scheme) == len(self.scheme_alt), (
            'Lists in variables ranges, scheme, and scheme_alt must '
            'all have the same length (see notes for details).')
        
        hex_pattern = r'^(?:[0-9a-fA-F]{3}){1,2}$'

        for scheme in [self.scheme, self.scheme_alt]:
            assert all(re.fullmatch(hex_pattern, h) for h in scheme), (
                f'Invalid hex code found:'
                f'{self.show_var("scheme", "scheme_alt")}'
                f'Did you make a typo?')

    def __init__(self, filepath: str):
        parser = configparser.ConfigParser()
        parser.read(filepath)

        # Flatted config dict
        self.__dict__ |= {k: v for d in parser.values() for k, v in d.items()}

        self.check_missing_vars()
        self.parse_dict()

        # Conditional validation
        try:
            self.validate()
        except AssertionError as e:
            Error(31.1).throw(*e.args)

    def check_missing_vars(self):
        if missing_vars := [
            v for v in ConfigVarTypes.__annotations__ if v not in self.__dict__
        ]:
            Error(30).throw(str(missing_vars))

    def parse_dict(self):
        for var, var_type in ConfigVarTypes.__annotations__.items():
            try:
                if var_type in [float, str]:
                    self.__dict__[var] = var_type(self.__dict__[var])
                
                elif var_type in [int, bool]:
                    # Fix str conversion issues such as '0' == True
                    self.__dict__[var] = var_type(float(self.__dict__[var]))

                else:  # Remaining is <class 'list'>
                    self.__dict__[var] = self.parse_list(
                        var_type.__args__[0],  # Extract the type of the list
                                               # ... e.g. <class 'float'> if
                                               # ... var_type is list[float]
                        self.__dict__[var])
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
                    f'{self.show_var(var)}'
                )

    T = TypeVar('T')

    def parse_list(self, var_type: Callable[[str], T], val: str) -> list[T]:
        raw_list = (
            val
            .replace('(', '')
            .replace(')', '')
            .split(','))

        return [var_type(v.strip()) for v in raw_list]

    def show_var(self, *vars: str) -> str:
        l = [f'    {var} = {self.__dict__[var]}' for var in vars]
        return '\n\n' + '\n'.join(l) + '\n'


if __name__ == '__main__':
    config = Config(abs_path('settings.ini')).__dict__
    print(config)
