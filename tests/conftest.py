from datetime import date

import pytest

from src.excel_orm import (
    Column,
    ExcelFile,
    PivotSheetSpec,
    SheetSpec,
    date_column,
    int_column,
    text_column,
)


@pytest.fixture
def models():
    class Car:
        make: Column[str] = text_column(header="Make", not_null=True)
        model: Column[str] = text_column(header="Model", not_null=True)
        year: Column[int] = int_column(header="Year", not_null=True)

    class ManufacturingPlant:
        name: Column[str] = text_column(header="Factory Name", not_null=True)
        location: Column[str] = text_column(header="Location")

    return Car, ManufacturingPlant


@pytest.fixture
def excel_file(models):
    Car, ManufacturingPlant = models
    sheet = SheetSpec(
        name="Cars",
        models=[Car, ManufacturingPlant],
        title_row=1,
        header_row=2,
        data_start_row=3,
        template_table_gap=2,
    )
    return ExcelFile(sheets=[sheet])


@pytest.fixture
def demand_model():
    class Demand:
        dt: Column[date] = date_column(header="Date")
        region: Column[str] = text_column(header="Region", not_null=True)
        value: Column[int] = int_column(header="Value", not_null=True)

    return Demand


@pytest.fixture
def pivot_excel_file(demand_model):
    Demand = demand_model

    # Pre-defined pivot values across the top of the sheet
    pivot_values = [
        date(2025, 6, 1),
        date(2025, 7, 1),
        date(2025, 8, 1),
    ]

    spec = PivotSheetSpec(
        name="Demand",
        model=Demand,
        pivot_field="dt",
        row_field="region",
        value_field="value",
        pivot_values=pivot_values,
        row_values=["NA", "EU"],  # seed some regions
    )

    return ExcelFile(sheets=[spec])
