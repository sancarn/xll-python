"""Worksheet functions copied into the disposable Excel smoke-test profile."""
import asyncio
from collections.abc import AsyncIterator
from itertools import count
from pathlib import Path
from typing import Annotated

from xll_python import ExcelError, Function, function


instances = count(1)


async def shared(key: str) -> Annotated[AsyncIterator[int], Function(name="PY.SHARED")]:
    instance = next(instances)
    try:
        yield instance
        await asyncio.Event().wait()
    finally:
        Path(__file__).with_name(f"closed-{instance}").touch()


@function(name="PY.PERCELL", generator_instances="many")
async def per_cell():
    async for value in shared(""):
        yield value


@function(name="PY.FINITE")
def finite(kind):
    if kind == "empty":
        return
    if kind == "error":
        raise RuntimeError("generator traceback sentinel")
    if kind == "array":
        yield [[1, 2], [3, 4]]
        return
    yield ExcelError(42)  # #N/A
