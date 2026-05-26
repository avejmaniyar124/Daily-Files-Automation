"""Upload NASU zip files to SharePoint. Run: py sharepoint_upload.py"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lapse.sharepoint import *  # noqa: F403
from lapse import sharepoint

if __name__ == "__main__":
    sys.exit(sharepoint.main())
