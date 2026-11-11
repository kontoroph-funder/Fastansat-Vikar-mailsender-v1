"""Confined temporary-directory cleanup, including long Windows profile paths."""
from contextlib import contextmanager
import os
from pathlib import Path
import shutil
import tempfile


def remove_tree(path: Path, within: Path) -> None:
    target, boundary = Path(path).resolve(), Path(within).resolve()
    if target == boundary or not target.is_relative_to(boundary):
        raise ValueError('Temporary cleanup escaped its parent directory')
    value = str(target)
    if os.name == 'nt' and not value.startswith('\\\\?\\'):
        value = ('\\\\?\\UNC\\' + value[2:]) if value.startswith('\\\\') else ('\\\\?\\' + value)
    shutil.rmtree(value)


@contextmanager
def render_directory(output_dir: Path):
    output_dir = Path(output_dir).resolve()
    temporary = Path(tempfile.mkdtemp(prefix='.render-', dir=output_dir))
    try:
        yield temporary
    finally:
        remove_tree(temporary, within=output_dir)
