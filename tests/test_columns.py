from __future__ import annotations

from datetime import date, datetime

import pytest

from src.excel_orm import Column, bool_column, date_column, int_column, text_column


def test_descriptor_stores_in_values_dict():
    class Foo:
        a: Column[str] = text_column(header="A")

    f = Foo()
    f._values = {}  # ty:ignore[unresolved-attribute]

    f.a = "hello"
    assert f._values["a"] == "hello"  # ty:ignore[unresolved-attribute]
    assert f.a == "hello"


def test_descriptor_class_level_access_returns_column():
    class Foo:
        a: Column[str] = text_column(header="A")

    assert isinstance(Foo.a, Column)


def test_not_null_validation_raises_on_none_or_empty():
    class Foo:
        a: Column[str] = text_column(header="A", not_null=True)

    f = Foo()
    f._values = {}  # ty:ignore[unresolved-attribute]

    with pytest.raises(ValueError):
        f.a = None  # type: ignore[assignment]

    with pytest.raises(ValueError):
        f.a = ""


def test_text_column_strip_and_none_to_empty_string():
    col = text_column(header="X", strip=True)
    assert col.parse_cell(None) == ""
    assert col.parse_cell("  hi  ") == "hi"
    assert col.parse_cell(123) == "123"


def test_text_column_no_strip():
    col = text_column(header="X", strip=False)
    assert col.parse_cell("  hi  ") == "  hi  "


def test_int_column_parsing():
    col = int_column(header="Y")
    assert col.parse_cell(None) == 0
    assert col.parse_cell("") == 0
    assert col.parse_cell("42") == 42
    assert col.parse_cell(7) == 7


def test_bool_column_parsing_valid_values():
    col = bool_column(header="B")
    assert col.parse_cell(None) is False
    assert col.parse_cell("") is False
    assert col.parse_cell(True) is True
    assert col.parse_cell("YES") is True
    assert col.parse_cell("0") is False
    assert col.parse_cell("n") is False


def test_bool_column_invalid_raises():
    col = bool_column(header="B")
    with pytest.raises(ValueError):
        col.parse_cell("maybe")


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("01-JUN-2025", date(2025, 6, 1)),
        ("01-jun-2025", date(2025, 6, 1)),  # case-insensitive month
        ("2025-06-01", date(2025, 6, 1)),
        ("2025/06/01", date(2025, 6, 1)),
        ("06/01/2025", date(2025, 6, 1)),  # per your formats list (US)
        (datetime(2025, 6, 1, 12, 30), date(2025, 6, 1)),
        (date(2025, 6, 1), date(2025, 6, 1)),
        ("2025-06-01T13:45:00", date(2025, 6, 1)),  # ISO datetime
    ],
)
def test_date_column_tryparse_cascade(raw, expected):
    col = date_column(header="D")
    assert col.parse_cell(raw) == expected


def test_date_column_empty_raises():
    col = date_column(header="D")
    with pytest.raises(ValueError):
        col.parse_cell(None)
    with pytest.raises(ValueError):
        col.parse_cell("")
    with pytest.raises(ValueError):
        col.parse_cell("   ")


def test_date_column_invalid_raises():
    col = date_column(header="D")
    with pytest.raises(ValueError):
        col.parse_cell("not-a-date")
