import configparser
from typing import TypedDict

from exceptions import Error
from utils import abs_path, iter_to_type


class ConfigVarTypes(TypedDict):
    update_check: bool
    ini_version_tag: str

    avatar_mode: bool
    last_clear_avatar_cache: int

    trigger_word: str
    ranges: list[float]
    scheme: list[str]
    scheme_alt: list[str]


def parse_list(config: str, val: str) -> list:
    print(ConfigVarTypes.__annotations__[config] == list)

    # Extract the type of the list
    # e.g. <class 'float'> if the config's type is list[float]
    list_type = ConfigVarTypes.__annotations__[config].__args__[0]

    return iter_to_type(
        list_type,
        [v.strip() for v in val.replace('(', '').replace(')', '').split(',')]
    )

# print(parse_list('avatar_mode', '(3, 5,24, 8)'))

if __name__ == '__main__':
    cfg_parser = configparser.ConfigParser()
    cfg_parser.read(abs_path('settings.ini'))

    config = {k: v for d in cfg_parser.values() for k, v in d.items()}

    print(config)

    if missing_vars := [
        v for v in ConfigVarTypes.__annotations__ if v not in config
    ]:
        Error(30).throw(str(missing_vars))

    # for cfg, cfg_type in ConfigVarTypes.__annotations__.values():
    #     if v in []
