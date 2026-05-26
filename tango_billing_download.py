"""
Download Tango PDF workflow billing files and create P{date}NASU_BILLING.zip.

Run from the project root:
  py tango_billing_download.py --billing-only --single-pass --no-wait --headed
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lapse.tango import *  # noqa: F403
from lapse import tango

if __name__ == "__main__":
    sys.exit(tango.main())
