from pathlib import Path
import sys

from rich.console import Console


REPO_URL = "https://github.com/SicariusBlack/mic-drop-results"
LATEST_RELEASE_URL = f"{REPO_URL}/releases/latest"
TEMPLATES_URL = f"{REPO_URL}/tree/main/templates"

if getattr(sys, "frozen", False):
    MAIN_DIR = Path(sys.executable).resolve().parent
else:
    MAIN_DIR = Path(__file__).resolve().parent

OUTPUT_DIR = MAIN_DIR / "output"
AVATAR_DIR = MAIN_DIR / "avatars"
TEMP_DIR = MAIN_DIR / ".temp"

console = Console(highlight=False)
padding = 4

# Mutable globals (usage: import constants; constants.var)
downloaded = 0
queue_len = 0
delay = 0
is_rate_limited = False
max_workers = 0
avatar_urls: list[tuple] = []
is_downloading = False
