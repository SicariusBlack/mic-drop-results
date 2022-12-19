from collections.abc import Callable
import configparser
import re
from typing import TypedDict, Any, TypeVar

from exceptions import Error
from utils import abs_path


class ConfigVarTypes(TypedDict):
    update_check: bool
    ini_version_tag: str

    avatar_mode: bool
    last_clear_avatar_cache: int

    trigger_word: str
    ranges: list[float]
    scheme: list[str]
    scheme_alt: list[str]


class ConfigVarConditions:
    def __init__(self, config) -> None:
        self.trigger_word: str
        self.ranges: list
        self.scheme: list
        self.scheme_alt: list

        self.__dict__.update(**config)

    def check(self):
        assert len(self.trigger_word) > 0, (
            'Variable trigger_word must not be empty.')
        
        assert len(self.ranges) == len(self.scheme) == len(self.scheme_alt), (
            'Lists in variables ranges, scheme, and scheme_alt must '
            'all have the same length (see notes for details).')
        
        hex_pattern = r'^(?:[0-9a-fA-F]{3}){1,2}$'

        for i, scheme in enumerate([self.scheme, self.scheme_alt]):
            assert all(re.fullmatch(hex_pattern, h) for h in scheme), (
                f'Invalid hex code found in\n'
                f'    {["scheme", "scheme_alt"][i]} = {scheme}\n'
                f'Did you make a typo?')


T = TypeVar('T')  # Pronounces "typed"


def parse_list(var_type: Callable[[str], T], val: str) -> list[T]:
    raw_list = (
        val
        .replace('(', '')
        .replace(')', '')
        .split(','))

    return [var_type(v.strip()) for v in raw_list]


def parse_config(config: dict[str, Any]) -> ConfigVarTypes:
    for var, var_type in ConfigVarTypes.__annotations__.items():
        try:
            if var_type in [float, str]:
                config[var] = var_type(config[var])
            
            elif var_type in [int, bool]:
                # Fix str conversion issues such as '0' == True
                config[var] = var_type(float(config[var]))

            else:  # Remaining is <class 'list'>
                config[var] = parse_list(
                    var_type.__args__[0], # Extract the type of the list
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

    return config  # type: ignore


def load_config(filepath: str) -> ConfigVarTypes:
    cfg_parser = configparser.ConfigParser()
    cfg_parser.read(filepath)

    config = {
        k: v for d in cfg_parser.values() for k, v in d.items()
    }

    if missing_vars := [
        v for v in ConfigVarTypes.__annotations__ if v not in config
    ]:
        Error(30).throw(str(missing_vars))

    config = parse_config(config)

    try:
        ConfigVarConditions(config).check()
    except AssertionError as e:
        Error(31.1).throw(*e.args)

    return config


if __name__ == '__main__':
    config = load_config(abs_path('settings.ini'))
