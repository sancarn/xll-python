import asyncio
import random
import time
from collections.abc import Iterator
from typing import Annotated
from xll_python import Function, function


@function(name="XLLPY.live_random", category="Python examples", generator_instances="many")
def live_random_number() -> Iterator[float]:
    """Yield a new random number every two seconds."""
    while True:
        yield random.random()
        time.sleep(2)


@function(name="XLLPY.scale", category="Python examples", thread_safe=True)
def scale(value: Annotated[float, "Number to scale"], factor: float = 2) -> float:
    """Multiply a number by a scale factor."""
    return value * factor


async def delayed(value: float, seconds: float = 2) -> Annotated[float, Function(name="XLLPY.delayed")]:
    """Return a number after the requested delay while the workbook remains usable."""
    await asyncio.sleep(seconds)
    return value


@function(name="XLLPY.transpose")
def transpose(values: list[list[object]]) -> list[list[object]]:
    """Transpose a rectangular range and spill the result."""
    return [list(column) for column in zip(*values)]

