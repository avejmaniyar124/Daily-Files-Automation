"""Slack alerts for missing Tango files. Run: py slack_notify.py --test"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lapse.slack import *  # noqa: F403
from lapse import slack

if __name__ == "__main__":
    sys.exit(slack.main())
