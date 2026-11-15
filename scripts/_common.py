"""Tiny helper shared by the scripts."""
import argparse
from pathlib import Path

from openpyxl import load_workbook


def existing_file(path):
    """argparse type: fail with a clear message (exit 2) if the file is missing or not a readable workbook."""
    if not Path(path).is_file():
        raise argparse.ArgumentTypeError(f"file not found: {path}")
    try:
        load_workbook(path, read_only=True).close()
    except Exception as e:
        raise argparse.ArgumentTypeError(f"cannot read {path} as an Excel workbook: {e}")
    return path
