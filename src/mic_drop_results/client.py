# Copyright 2023 Phan Huy

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.

from concurrent.futures import ThreadPoolExecutor
from enum import Enum
import requests
import time
from functools import partial
from threading import Event
from typing import Iterable
from urllib.request import Request, urlopen
from urllib.error import URLError

import cv2
import numpy as np
from rich.progress import (
    BarColumn,
    DownloadColumn,
    Progress,
    TaskID,
    TextColumn,
    TimeRemainingColumn,
    TransferSpeedColumn,
)

import constants
from constants import *
from exceptions import *
from utils import is_number, get_avatar_dir


done_event = Event()


# Section A: GitHub API
class ProgramStatus(Enum):
    UPDATE_AVAILABLE = "update available"
    UP_TO_DATE = "latest"
    BETA = "beta"


def fetch_latest_version() -> tuple[str, str]:
    response = requests.get(
        "https://api.github.com/repos/SicariusBlack/mic-drop-results/releases/latest",
        timeout=3,
    )
    tag = response.json()["tag_name"].lstrip("v")
    summary = response.json()["body"].partition("\n")[0].lstrip("# ")
    return (tag, summary)


# Section B: Discord's API
def _fetch_avatar_url(uid: str, api_token: str) -> str | None:  # TODO: docstring
    if not is_number(uid):
        return None

    # Try sending out a request to the API for the avatar's hash
    try:
        header = {"Authorization": f"Bot {api_token}"}
        response = requests.get(
            f"https://discord.com/api/v10/users/{uid}", headers=header, timeout=20
        )
        # if not ('message' in response.json() or 'avatar' in response.json())
    except (requests.exceptions.ConnectionError, requests.exceptions.ReadTimeout) as e:
        raise ConnectionError from e

    # Try extracting the hash and return the complete link if succeed
    try:
        if response.json()["avatar"] is not None:
            return "https://cdn.discordapp.com/avatars/{}/{}.png".format(
                uid, response.json()["avatar"]
            )

        if response.json()["discriminator"] == "0000":
            return None
        # Return default avatar
        # https://discord.com/developers/docs/reference#image-formatting-cdn-endpoints
        return "https://cdn.discordapp.com/embed/avatars/{}.png".format(
            int(response.json()["discriminator"]) % 5
        )
    except KeyError as e:
        msg = response.json()["message"].lower()
        if "401: unauthorized" in msg:  # invalid token
            raise InvalidTokenError(api_token) from e

        elif "limit" in msg:
            r = response.json()["retry_after"] + 10
            console.print("\033[A\033[2K" f"You are being rate-limited by the API.")
            time.sleep(r)
            _fetch_avatar_url(uid, api_token)

        elif "unknown" not in msg:
            raise DiscordAPIError(api_token, response.json()) from e


def _download(avatar_url: str, img_dir: Path) -> None:
    try:
        req = urlopen(Request(avatar_url, headers={"User-Agent": "Mozilla/5.0"}))
        arr = np.asarray(bytearray(req.read()), dtype=np.uint8)
        img = cv2.imdecode(arr, -1)

        cv2.imwrite(str(img_dir), img)
    except URLError as e:
        raise ConnectionError from e


def fetch_avatar(uid, api_token, size, progress, task):
    if avatar_url := _fetch_avatar_url(uid, api_token):
        progress.update(
            task,
            advance=1,
            description=_get_download_banner(avatar_url),
        )
        avatar_url += f"?size={size}"
        constants.avatar_urls.append((uid, avatar_url))


def download_avatars():
    while constants.is_downloading:
        with ThreadPoolExecutor(max_workers=4) as pool:
            while len(constants.avatar_urls) > 0:
                uid, avatar_url = constants.avatar_urls[0]
                img_dir = get_avatar_dir(uid)

                pool.submit(_download, avatar_url, img_dir)

                constants.avatar_urls.pop(0)


def _get_download_banner(desc: str) -> str:
    return f"Downloading avatars... ({constants.queue_len} in queue)\n{desc}"
