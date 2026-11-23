#!/usr/bin/env python3
"""Resolve the external Archery capability without machine-specific paths."""
import os
from pathlib import Path

DEFAULT_ARCHERY_DIR = (
    Path.home()
    / "ai-workspace"
    / "shared"
    / "benlai-trace-debug"
    / "scripts"
    / "archery"
)

ARCHERY_DIR = Path(
    os.environ.get("BENLAI_ARCHERY_DIR", str(DEFAULT_ARCHERY_DIR))
).expanduser()
