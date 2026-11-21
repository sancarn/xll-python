"""Public API for worksheet function modules loaded by the Python ribbon."""
from dataclasses import dataclass
from typing import Callable, Literal, TypeVar, TypedDict, Unpack, overload

from _xll_application import Application


@dataclass(frozen=True)
class Function:
    """Excel registration options, usable in Annotated[return_type, Function(...)]."""
    name: str = ""
    category: str = "Python"
    description: str = ""
    asynchronous: bool | None = None
    volatile: bool = False
    thread_safe: bool = False
    hidden: bool = False
    macro_type: bool = False
    cluster_safe: bool = False
    help_topic: str = ""
    generator_instances: Literal["one", "many"] = "one"


class _Options(TypedDict, total=False):
    name: str
    category: str
    description: str
    asynchronous: bool | None
    volatile: bool
    thread_safe: bool
    hidden: bool
    macro_type: bool
    cluster_safe: bool
    help_topic: str
    generator_instances: Literal["one", "many"]


_F = TypeVar("_F", bound=Callable)


@overload
def function(fn: _F, **options: Unpack[_Options]) -> _F: ...


@overload
def function(fn: None = None, **options: Unpack[_Options]) -> Callable[[_F], _F]: ...


def function(fn=None, **options):
    """Export a worksheet function. Generators publish each yield through Excel RTD."""
    def decorate(target):
        target.__excel_function__ = Function(**options)
        return target
    return decorate if fn is None else decorate(fn)


@overload
def macro(fn: _F, *, name: str = "") -> _F: ...


@overload
def macro(fn: None = None, *, name: str = "") -> Callable[[_F], _F]: ...


def macro(fn=None, *, name=""):
    """Expose a command callable without arguments in Excel's macro pickers."""
    def decorate(target):
        target.__excel_macro__ = name
        return target
    return decorate if fn is None else decorate(fn)


@dataclass(frozen=True)
class ExcelError:
    """An Excel error value, using its Excel C API error code."""
    code: int


@dataclass(frozen=True)
class Reference:
    """A worksheet reference. Areas are inclusive (first row, last row, first column, last column), zero-based."""
    sheet_id: int
    areas: tuple[tuple[int, int, int, int], ...]

