from enum import Enum
import requests
import time
from urllib.request import Request, urlopen

import cv2
import numpy as np

from exceptions import Error, ErrorType
from utils import is_number


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


# Section B: Discord API
def fetch_avatar_url(id: str, api_token: str) -> str | None:
    if not is_number(id):
        return None

    # Try sending out a request to the API for the avatar's hash
    try:
        header = {'Authorization': f'Bot {api_token}'}
        response = requests.get(
            f'https://discord.com/api/v9/users/{id}', headers=header
        )
    except requests.exceptions.ConnectionError:
        Error(20).throw(err_type=ErrorType.WARNING)
        return None

    # Try extracting the hash and return the complete link if succeed
    try:
        return 'https://cdn.discordapp.com/avatars/{}/{}.png'.format(
            id, response.json()['avatar'])
    except KeyError as e:
        # Invalid token or a user account has been deleted (hypothesis)
        # TODO: Test out the hypothesis
        if '401: unauthorized' in response.json()['message'].lower():
            Error(21.1).throw(api_token, response.json())

        elif 'rate-limit' in response.json()['message'].lower():
            time.sleep(response.json()['retry_after'])
            fetch_avatar_url(id, api_token)

        else:
            raise response.json() from e


def download_avatar(uid, avatar_dir, api_token):
    uid = uid.strip().replace('_', '')
    img_path = f'{avatar_dir}_{uid.strip()}.png'

    if avatar_url := fetch_avatar_url(uid, api_token):
        req = urlopen(Request(
            avatar_url, headers={'User-Agent': 'Mozilla/5.0'}))
        arr = np.asarray(bytearray(req.read()), dtype=np.uint8)
        img = cv2.imdecode(arr, -1)

        cv2.imwrite(img_path, img)
