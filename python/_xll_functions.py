"""Worksheet function registration, argument conversion and return values."""
import asyncio
import dataclasses
import datetime
import inspect
import math
import re
import typing

from xll_python import ExcelError, Function, Reference


def _command_label(fn):
    if inspect.isgeneratorfunction(fn) or inspect.isasyncgenfunction(fn):
        raise TypeError("a macro cannot be a generator")
    try:
        inspect.signature(fn).bind()
    except TypeError as error:
        raise TypeError("a macro must be callable without arguments") from error
    name = getattr(fn, "__excel_macro__", "")
    if not isinstance(name, str):
        raise TypeError("a macro name must be a string")
    return name.strip() or fn.__name__


def _run_command(fn):
    result = fn()
    if inspect.isawaitable(result):
        asyncio.run(result)


def _annotation(annotation):
    if typing.get_origin(annotation) is typing.Annotated:
        return typing.get_args(annotation)
    return (annotation,)


def _cell_name(name):
    if re.fullmatch(r"R(?:[1-9][0-9]*)?C(?:[1-9][0-9]*)?|R|C", name, re.I):
        return True
    match = re.fullmatch(r"([A-Z]{1,3})([1-9][0-9]{0,6})", name.upper())
    if match is None:
        return False
    column = 0
    for letter in match[1]:
        column = 26 * column + ord(letter) - ord("A") + 1
    return column <= 16384 and int(match[2]) <= 1048576


def _register(fn, functions, entries, console_name=None):
    opt = getattr(fn, "__excel_function__", None)
    raw_return = fn.__annotations__.get("return")
    if console_name is None and opt is None and not (typing.get_origin(raw_return) is typing.Annotated or
                            isinstance(raw_return, str) and "Annotated" in raw_return):
        return
    hints = typing.get_type_hints(fn, include_extras=True)
    ret = _annotation(hints.get("return"))
    opt = opt or next((x for x in ret[1:] if isinstance(x, Function)), None)
    if console_name is not None and opt is None:
        opt = Function(category="Python Console", volatile=True)
    if opt is None:
        return
    name = console_name or opt.name or fn.__name__
    if (not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]{0,254}", name)
        or _cell_name(name)):
        raise ValueError(f"Invalid or cell-like Excel function name: {name}")
    if name.upper() in functions:
        raise ValueError(f"Duplicate worksheet function: {name}")
    signature = inspect.signature(fn)
    args = []
    for par in signature.parameters.values():
        if par.kind not in (par.POSITIONAL_ONLY, par.POSITIONAL_OR_KEYWORD):
            raise TypeError(f"{name} needs a fixed positional signature")
        meta = _annotation(hints.get(par.name))[1:]
        args.append(dict(name=par.name, reference=_annotation(hints.get(par.name))[0] is Reference,
                         description=next((m for m in meta if isinstance(m, str)), "")))
    generator = inspect.isgeneratorfunction(fn) or inspect.isasyncgenfunction(fn)
    if opt.generator_instances not in ("one", "many"):
        raise ValueError(f"{name}: generator_instances must be 'one' or 'many'")
    if not generator and opt.generator_instances == "many":
        raise ValueError(f"{name}: generator_instances='many' requires a generator function")
    if generator and opt.asynchronous is False:
        raise ValueError(f"{name}: generators require asynchronous registration")
    asynchronous = generator or (opt.asynchronous if opt.asynchronous is not None else inspect.iscoroutinefunction(fn))
    if inspect.iscoroutinefunction(fn) and not asynchronous:
        raise ValueError(f"{name}: async def requires asynchronous registration")
    if opt.macro_type and (opt.thread_safe or asynchronous):
        raise ValueError(f"{name}: incompatible Excel macro-type registration flags")
    if len(args) > (254 if asynchronous else 255):
        raise ValueError(f"{name}: too many Excel arguments")
    entries.append(dict(dataclasses.asdict(opt), name=name, arguments=args,
                        generator=generator, asynchronous=asynchronous, description=opt.description or inspect.getdoc(fn) or ""))
    functions[name.upper()] = (fn, signature, hints)


def _decode(value):
    if isinstance(value, list):
        return [_decode(x) for x in value]
    if isinstance(value, dict) and "error" in value:
        return ExcelError(value["error"])
    if isinstance(value, dict) and "areas" in value:
        return Reference(value["sheet_id"], tuple(tuple(a) for a in value["areas"]))
    return value


def _convert(value, annotation):
    typ = _annotation(annotation)[0]
    if value is None or isinstance(value, ExcelError):
        return value
    if typing.get_origin(typ) is list:
        arg = typing.get_args(typ)[0]
        if not isinstance(value, list):
            value = [[value]] if typing.get_origin(arg) is list else [value]
        return [_convert(x, arg) for x in value]
    if typ in (int, float, bool, str):
        return typ(value)
    if typ in (datetime.datetime, datetime.date):
        dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=value)
        return dt if typ is datetime.datetime else dt.date()
    return value


def _encode(value):
    if isinstance(value, ExcelError):
        return {"error": value.code}
    if isinstance(value, (datetime.date, datetime.datetime)):
        if not isinstance(value, datetime.datetime):
            value = datetime.datetime.combine(value, datetime.time())
        return (value - datetime.datetime(1899, 12, 30)).total_seconds() / 86400
    if isinstance(value, (list, tuple)):
        return [_encode(x) for x in value]
    if isinstance(value, float) and not math.isfinite(value):
        return {"error": 36}
    if hasattr(value, "tolist"):
        return _encode(value.tolist())
    return value


def _call(registration, args):
    fn, signature, hints = registration
    values = []
    for par, value in zip(signature.parameters.values(), args):
        if isinstance(value, dict) and value.get("missing") is True:
            if par.default is par.empty:
                raise TypeError(f"Missing argument: {par.name}")
            values.append(par.default)
        else:
            values.append(_convert(_decode(value), hints.get(par.name)))
    return fn(*values)


def _invoke(registration, args):
    result = _call(registration, args)
    if inspect.isawaitable(result):
        result = asyncio.run(result)
    return _encode(result)
