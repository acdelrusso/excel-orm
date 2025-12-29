from __future__ import annotations

from src.excel_orm import Column, int_column, text_column
from src.excel_orm.orm import (
    _camel_to_snake,
    _display_name_for_model,
    _get_model_columns,
    _instantiate_model,
    _pluralize,
    _repo_name_for_model,
)


def test_camel_to_snake():
    assert _camel_to_snake("Car") == "car"
    assert _camel_to_snake("ManufacturingPlant") == "manufacturing_plant"
    assert _camel_to_snake("HTTPServer") == "http_server"


def test_pluralize():
    assert _pluralize("car") == "cars"
    assert _pluralize("class") == "class"  # your simple rule: if endswith s, keep as-is


def test_repo_and_display_name():
    class ManufacturingPlant:
        pass

    assert _repo_name_for_model(ManufacturingPlant) == "manufacturing_plants"
    assert _display_name_for_model(ManufacturingPlant) == "Manufacturing Plants"


def test_get_model_columns_respects_annotation_order():
    class A:
        x: Column[str] = text_column(header="X")
        y: Column[int] = int_column(header="Y")

    cols = _get_model_columns(A)
    assert [c.name for c in cols] == ["x", "y"]


def test_instantiate_model_sets_defaults_and_requires_set_name():
    class A:
        x: Column[str] = text_column(header="X", default="dflt")

    a = _instantiate_model(A)
    assert a._values["x"] == "dflt"  # ty:ignore[unresolved-attribute]
