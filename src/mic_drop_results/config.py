from collections.abc import Callable
import configparser
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


T = TypeVar('T')  # Pronounces "typed"


def parse_list(var_type: Callable[[str], T], val: str) -> list[T]:
    raw_list = val.replace('(', '').replace(')', '').split(',')
    return [var_type(v.strip()) for v in raw_list]


if __name__ == '__main__':
    cfg_parser = configparser.ConfigParser()
    cfg_parser.read(abs_path('settings.ini'))

    config: dict[str, Any] = {
        k: v for d in cfg_parser.values() for k, v in d.items()
    }

    if missing_vars := [
        v for v in ConfigVarTypes.__annotations__ if v not in config
    ]:
        Error(30).throw(str(missing_vars))

    for var, var_type in ConfigVarTypes.__annotations__.items():
        if var_type in [float, int, bool, str]:
            config[var] = var_type(config[var])

        else:  # Remaining is <class 'list'>
            config[var] = parse_list(
                var_type.__args__[0],  # Extract the type of the list
                                       # ... e.g. <class 'float'> if
                                       # ... var_type is list[float]
                config[var]
            )

    print(config)
