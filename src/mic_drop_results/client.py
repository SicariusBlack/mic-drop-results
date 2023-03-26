# Copyright 2023 Phan Huy

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.

from enum import Enum
import requests
import time
from urllib.request import Request, urlopen
from urllib.error import URLError

import cv2
import numpy as np

from constants import *
from exceptions import *
from utils import is_number, get_avatar_path


# Section A: GitHub API
class ProgramStatus(Enum):
    UPDATE_AVAILABLE = 'update available'
    UP_TO_DATE = 'latest'
    BETA = 'beta'


def fetch_latest_version() -> tuple[str, str]:
    response = requests.get(
        'https://api.github.com/repos/banz04/mic-drop-results/releases/latest',
        timeout=3,
    )
    tag = response.json()['tag_name'].lstrip('v')
    summary = response.json()['body'].partition('\n')[0]
    return (tag, summary)


# Section B: Discord's API
def fetch_avatar_url(uid: str, api_token: str) -> str | None:  # TODO: docstring
    if not is_number(uid):
        return None

    time.sleep(0.02)
    # Try sending out a request to the API for the avatar's hash
    try:
        header = {'Authorization': f'Bot {api_token}'}
        response = requests.get(
            f'https://discord.com/api/v10/users/{uid}', headers=header,
            timeout=10
        )
        #if not ('message' in response.json() or 'avatar' in response.json())
    except (requests.exceptions.ConnectionError,
            requests.exceptions.ReadTimeout) as e:
        raise ConnectionError from e

    # Try extracting the hash and return the complete link if succeed
    try:
        if response.json()['avatar'] is not None:
            return 'https://cdn.discordapp.com/avatars/{}/{}.png'.format(
                uid, response.json()['avatar'])

        if response.json()['discriminator'] == '0000':
            return None
        # Return default avatar
        # https://discord.com/developers/docs/reference#image-formatting-cdn-endpoints
        return 'https://cdn.discordapp.com/embed/avatars/{}.png'.format(
            int(response.json()['discriminator']) % 5)
    except KeyError as e:
        msg = response.json()['message'].lower()
        if '401: unauthorized' in msg:  # invalid token
            raise InvalidTokenError(api_token, response.json()) from e

        elif 'limit' in msg:
            r = response.json()['retry_after'] + 10
            print(
                '\033[A\033[2K'
                f'You are being rate-limited by the API.')
            time.sleep(r)
            fetch_avatar_url(uid, api_token)

        elif 'unknown' not in msg:
            raise DiscordAPIError(api_token, response.json()) from e


def download_avatar(uid: str, api_token: str, size: int) -> None:
    img_path = get_avatar_path(uid)

    try:
        if avatar_url := fetch_avatar_url(uid, api_token):
            print('\033[A\033[2K' + avatar_url)
            avatar_url += f'?size={size}'
            req = urlopen(Request(
                avatar_url, headers={'User-Agent': 'Mozilla/5.0'}))
            arr = np.asarray(bytearray(req.read()), dtype=np.uint8)
            img = cv2.imdecode(arr, -1)

            cv2.imwrite(str(img_path), img)
    except URLError as e:
        raise ConnectionError from e
