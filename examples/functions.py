import asyncio
from typing import Annotated
from xll_python import Function, function


@function(name="PY.SCALE", category="Python examples", thread_safe=True)
def scale(value: Annotated[float, "Number to scale"], factor: float = 2) -> float:
    """Multiply a number by a scale factor."""
    return value * factor


async def delayed(value: float) -> Annotated[float, Function(name="PY.DELAYED")]:
    """Return a number asynchronously."""
    await asyncio.sleep(0.1)
    return value


@function(name="PY.TRANSPOSE")
def transpose(values: list[list[object]]) -> list[list[object]]:
    """Transpose a rectangular range and spill the result."""
    return [list(column) for column in zip(*values)]

