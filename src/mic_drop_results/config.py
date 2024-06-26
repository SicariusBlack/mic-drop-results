# Copyright 2023 Phan Huy

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.

from collections.abc import Callable
import configparser
from typing import Any

from compiled_regex import *
from errors import Error


class ConfigVarTypes:
    update_check: bool
    avatar_mode: bool
    statistics: bool
    avatar_resolution: int

    sort_orders: list[bool]

    trigger_word: str
    ranges: list[float]
    scheme: list[str]
    scheme_alt: list[str]


class Config(ConfigVarTypes):  # TODO: add docstrings
    def __init__(self, file_dir: str):
        parser = configparser.ConfigParser()
        parser.read(file_dir)

        # Flatten config dict
        self.config: dict[str, Any] = {
            k: v for d in parser.values() for k, v in d.items()
        }

        self._check_missing_vars()
        self._parse_config()

        try:
            self._validate(self.config)  # validate config vars' conditions
        except AssertionError as e:
            Error(31.1).throw("[red]" + e.args[0] + "[/red]")

        self.__dict__ = self.config  # assign config vars to class attributes

    def _validate(self, cfg: dict[str, Any]) -> None:
        resolution_presets = [16, 32, 64, 80, 100, 128, 256, 512, 1024, 2048]
        assert (
            cfg["avatar_resolution"] in resolution_presets
        ), "Avatar resolution must be taken from the list of available resolutions."

        assert (
            len(cfg["trigger_word"]) > 0
        ), 'Config variable "trigger_word" cannot be empty.'

        assert (
            len(cfg["ranges"]) == len(cfg["scheme"]) == len(cfg["scheme_alt"])
        ), 'The "ranges", "scheme", and "scheme_alt" lists must all have the same, matching length.'

        assert all(
            match_hex.fullmatch(h)
            for scheme in [cfg["scheme"], cfg["scheme_alt"]]
            for h in scheme
        ), "Invalid hex color code found in:" + self._show_var("scheme", "scheme_alt")

    def _check_missing_vars(self) -> None:
        if missing_vars := [
            v for v in ConfigVarTypes.__annotations__ if v not in self.config
        ]:
            Error(30).throw(str(missing_vars))

    def _parse_config(self) -> None:  # sourcery skip: list-literal
        for name, var_type in ConfigVarTypes.__annotations__.items():
            try:
                match var_type():
                    case str():
                        self.config[name] = str(self.config[name]).strip('"')

                    case float():
                        self.config[name] = float(self.config[name])

                    case int() | bool():
                        # Fix conversion issues such as '0' == True
                        self.config[name] = var_type(float(self.config[name]))

                    case list():
                        self.config[name] = self._parse_list(
                            var_type, self.config[name]
                        )

            except ValueError:
                if var_type() == list():
                    type_name = f"list of {var_type.__args__[0].__name__}"
                else:
                    type_name = var_type.__name__

                Error(31).throw(
                    f"Failed to convert the following config into type:  <{type_name}>"
                    + self._show_var(name)
                )

    def _parse_list(self, list_type: Callable[[str], Any], val: str) -> list:
        ele_type = list_type.__args__[
            0
        ]  # extract the types of individual elements, e.g. <class 'float'> if list_type = list[float]
        raw_list = val.strip("[]").split(",")

        match ele_type():
            case bool():
                ele_type = lambda v: bool(float(v))
            case int():
                ele_type = lambda v: int(float(v))

        return [ele_type(element.strip()) for element in raw_list]

    def _show_var(self, *vars: str) -> str:  # TODO
        l = [f"    {name} = {self.config[name]}" for name in vars]
        return "\n\n[red]" + "\n".join(l) + "[/red]"
