"""Заполните .env и нажмите Run в PyCharm."""

import os
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from yandex_attribution.cli import env_arguments, main  # noqa: E402


if __name__ == "__main__":
    os.chdir(PROJECT_ROOT)
    main(env_arguments())
