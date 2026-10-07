"""成果物を同一ファイルシステム上で原子的に確定する小さなI/O層。"""

from __future__ import annotations

import copy
from contextlib import contextmanager
import json
import os
from pathlib import Path
import tempfile
from typing import Callable, Iterator, TypeVar


FrozenValue = TypeVar("FrozenValue")


@contextmanager
def atomic_output_path(path: Path) -> Iterator[Path]:
    """同一ディレクトリの一時パスを返し、正常終了時だけ置換する。"""
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{destination.name}.",
        suffix=".tmp",
        dir=destination.parent,
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        yield temporary
        if not temporary.is_file():
            raise FileNotFoundError(f"一時成果物が作成されませんでした: {temporary}")
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


def atomic_write_bytes(path: Path, data: bytes) -> None:
    with atomic_output_path(path) as temporary:
        with temporary.open("wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())


def atomic_write_text(path: Path, text: str, *, encoding: str = "utf-8") -> None:
    atomic_write_bytes(Path(path), text.encode(encoding))


def atomic_write_json(path: Path, value: object) -> None:
    atomic_write_text(
        Path(path),
        json.dumps(value, ensure_ascii=False, indent=2, allow_nan=False),
    )


def atomic_render_immutable(
        path: Path,
        frozen_value: FrozenValue,
        renderer: Callable[[Path], None],
        *,
        mutation_message: str) -> None:
    """Render atomically and commit only when the approved value stayed immutable.

    ``atomic_output_path`` alone protects against a partially written file, but a
    renderer can still mutate the in-memory review object before the temporary
    artifact is committed.  Keeping the comparison inside the context prevents
    a failed render from replacing a previously approved artifact.
    """
    approved_value = copy.deepcopy(frozen_value)
    with atomic_output_path(Path(path)) as temporary:
        renderer(temporary)
        if frozen_value != approved_value:
            raise RuntimeError(mutation_message)
