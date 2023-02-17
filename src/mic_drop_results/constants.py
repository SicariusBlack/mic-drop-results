from pathlib import Path

REPO_URL = 'https://github.com/banz04/mic-drop-results'
LATEST_RELEASE_URL = f'{REPO_URL}/releases/latest'
TEMPLATES_URL = f'{REPO_URL}/tree/main/templates'

MAIN_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = MAIN_DIR / 'output'
AVATAR_DIR = MAIN_DIR / 'avatars'
TEMP_DIR = MAIN_DIR / '.temp'
