import pytest

from src.column import Column, int_column, text_column
from src.orm import ExcelFile, SheetSpec


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
