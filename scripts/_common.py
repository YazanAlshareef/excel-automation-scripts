"""Tiny helper shared by the scripts."""
import argparse
from pathlib import Path


def existing_file(path):
    """argparse type: fail with a clear message (exit 2) if the file is missing."""
    if not Path(path).is_file():
        raise argparse.ArgumentTypeError(f"file not found: {path}")
    return path
