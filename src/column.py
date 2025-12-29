from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, TypeVar

T = TypeVar("T")


@dataclass(frozen=True)
class ColumnSpec[T]:
    header: str | None = None  # header string in Excel
    index: int | None = None  # 0-based index fallback (optional)
    default: T | None = None
    required: bool = False  # must be present in sheet mapping
    not_null: bool = False  # parsed value cannot be None/empty
    excludes: set[Any] | None = None  # raw values that mark row as excluded
    parser: Callable[[Any], T] = lambda x: x  # raw -> parsed
    renderer: Callable[[T | None], Any] = lambda x: x  # parsed -> raw
    validator: Callable[[T | None], None] = lambda _: None


class Column[T]:
    def __init__(self, spec: ColumnSpec[T]):
        self.spec = spec

    def __set_name__(self, owner, name: str):
        self.name = name

    def __get__(self, obj, objtype=None) -> T | Column | None:
        if obj is None:
            return self
        return obj._values.get(self.name)  # centralized storage

    def __set__(self, obj, value: T | None):
        self.validate(value)
        obj._values[self.name] = value

    def parse_cell(self, raw: Any) -> T | None:
        return self.spec.parser(raw)

    def render_cell(self, value: T | None) -> Any:
        return self.spec.renderer(value)

    def validate(self, value: T | None) -> None:
        if self.spec.not_null and (value is None or value == ""):
            raise ValueError(f"{self.name} cannot be null/empty")
        self.spec.validator(value)


def text_column(
    header: str | None = None,
    *,
    index: int | None = None,
    default: str | None = None,
    strip: bool = True,
    not_null: bool = False,
):
    def parse(raw: Any) -> str:
        if raw is None:
            return ""
        s = str(raw)
        if strip:
            s = s.strip()
        return s

    return Column(
        ColumnSpec[str](
            header=header,
            index=index,
            default=default,
            not_null=not_null,
            parser=parse,
            renderer=lambda v: "" if v is None else v,
        )
    )


def int_column(
    header: str | None = None,
    *,
    index: int | None = None,
    default: int | None = None,
    not_null: bool = False,
):
    def parse(raw: Any) -> int:
        if raw is None or raw == "":
            return 0
        return int(raw)

    return Column(
        ColumnSpec[int](
            header=header,
            index=index,
            default=default,
            not_null=not_null,
            parser=parse,
        )
    )


def bool_column(
    header: str | None = None, *, index: int | None = None, default: bool | None = None
):
    def parse(raw: Any) -> bool:
        if raw is None or raw == "":
            return False
        if isinstance(raw, bool):
            return raw
        s = str(raw).strip().lower()
        if s in {"true", "t", "yes", "y", "1"}:
            return True
        if s in {"false", "f", "no", "n", "0"}:
            return False
        raise ValueError(f"Invalid boolean: {raw}")

    return Column(
        ColumnSpec[bool](
            header=header,
            index=index,
            default=default,
            parser=parse,
        )
    )


def date_column(
    header: str | None = None, *, index: int | None = None, default: date | None = None
):
    _DATE_FORMATS: tuple[str, ...] = (
        "%d-%b-%Y",  # 01-JUN-2025  (your requirement)
        "%d-%b-%y",  # 01-JUN-25
        "%d %b %Y",  # 01 JUN 2025
        "%d %b %y",  # 01 JUN 25
        "%d/%b/%Y",  # 01/JUN/2025
        "%Y-%m-%d",  # 2025-06-01
        "%Y/%m/%d",  # 2025/06/01
        "%m/%d/%Y",  # 06/01/2025
        "%m/%d/%y",  # 06/01/25
        "%d/%m/%Y",  # 01/06/2025
        "%d/%m/%y",  # 01/06/25
    )

    def parse(raw: Any) -> date:
        if raw is None or raw == "":
            raise ValueError("Date Value was empty")

        if isinstance(raw, date) and not isinstance(raw, datetime):
            return raw

        if isinstance(raw, datetime):
            return raw.date()

        s = str(raw).strip()
        if s == "":
            raise ValueError("Date Value was empty")

        # 1) ISO-8601 fast path (handles "2025-06-01" and "2025-06-01T13:45:00", etc.)
        try:
            dt = datetime.fromisoformat(s)
            return dt.date()
        except ValueError:
            pass

        # 2) Try known patterns (case-insensitive month abbreviations like JUN)
        s_norm = s.upper()

        for fmt in _DATE_FORMATS:
            try:
                return datetime.strptime(s_norm, fmt).date()
            except ValueError:
                continue

        raise ValueError(f"Invalid date value: {raw!r}")

    return Column(
        ColumnSpec[date](
            header=header,
            index=index,
            default=default,
            parser=parse,
            renderer=lambda d: None if d is None else d,  # openpyxl handles date types
        )
    )
