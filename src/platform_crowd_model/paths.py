"""Where the repo keeps its data."""

from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
"""The repo's root."""

DATA_DIR = REPO / "data"
"""The repo's `data/`, with each platform's VCEs and dimensions."""

CACHE_DIR = REPO / ".cache"
"""Where the `data` commands cache the sources they download."""
