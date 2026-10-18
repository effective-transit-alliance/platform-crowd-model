"""Tests that `.cargo/config.toml` matches the environment `maturin` builds `pyo3` in."""

import struct
import sys
import tomllib

from platform_crowd_model.paths import REPO


def test_pyo3_environment_signature_matches_python() -> None:
    """
    `.cargo/config.toml`'s `$PYO3_ENVIRONMENT_SIGNATURE` is the one `maturin` gives `pyo3`
    for this Python, or else plain `cargo` and `maturin` rebuild `pyo3` for each other.
    """
    with (REPO / ".cargo" / "config.toml").open("rb") as f:
        signature = tomllib.load(f)["env"]["PYO3_ENVIRONMENT_SIGNATURE"]
    version = sys.version_info
    pointer_width = struct.calcsize("P") * 8
    expected = f"{sys.implementation.name}-{version.major}.{version.minor}-{pointer_width}bit"
    assert signature == expected, (
        f"update `$PYO3_ENVIRONMENT_SIGNATURE` in `.cargo/config.toml` to {expected!r}"
    )
