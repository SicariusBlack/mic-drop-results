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


T = TypeVar('T')  # Pronounces "typed"


class Config(ConfigVarTypes):
    def __init__(self, filepath: str):
        cfg_parser = configparser.ConfigParser()
        cfg_parser.read(filepath)

        config = {
            k: v for d in cfg_parser.values() for k, v in d.items()
        }

        if missing_vars := [
            v for v in ConfigVarTypes.__annotations__ if v not in config
        ]:
            Error(30).throw(str(missing_vars))

        config = self.parse_config(config)

        try:
            self.check()
        except AssertionError as e:
            Error(31.1).throw(*e.args)

        return config


    def parse_list(self, var_type: Callable[[str], T], val: str) -> list[T]:
        raw_list = (
            val
            .replace('(', '')
            .replace(')', '')
            .split(','))

        return [var_type(v.strip()) for v in raw_list]


    def parse_config(self, config: dict[str, Any]):
        for var, var_type in ConfigVarTypes.__annotations__.items():
            try:
                if var_type in [float, str]:
                    config[var] = var_type(config[var])
                
                elif var_type in [int, bool]:
                    # Fix str conversion issues such as '0' == True
                    config[var] = var_type(float(config[var]))

                else:  # Remaining is <class 'list'>
                    config[var] = self.parse_list(
                        var_type.__args__[0],  # Extract the type of the list
                                               # ... e.g. <class 'float'> if
                                               # ... var_type is list[float]
                        config[var])
            except ValueError:
                if var_type.__name__ == 'list':
                    type_name = f'list of {var_type.__args__[0].__name__}'
                    # e.g. 'list of float'
                else:
                    type_name = var_type.__name__
                    # e.g. 'float'

                Error(31).throw(
                    f'Failed to convert the following '
                    f'variable to type: <{type_name}>',
                    f'    {var} = {config[var]}'
                )

        self.__dict__.update(**config)

    def check(self):
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
    
    def show_var(self, *vars: str) -> str:
        l = [f'    {var} = {self.__dict__[var]}' for var in vars]
        return '\n\n' + '\n'.join(l) + '\n\n'


if __name__ == '__main__':
    config = Config(abs_path('settings.ini'))
