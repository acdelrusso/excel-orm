# Excel ORM

A lightweight, typed “Excel ORM” for generating Excel templates and parsing Excel workbooks into Python objects using column descriptors.

This project is designed for the common enterprise pattern where you:
1) generate a structured `.xlsx` template for users,
2) let users fill it in,
3) load the workbook back into Python, producing typed objects grouped by model.

It uses `openpyxl` for reading/writing Excel files and supports multiple model “tables” on the same worksheet.

---

## Features

- **Typed column descriptors** (`text_column`, `int_column`, `bool_column`, `date_column`)
- **Template generation** with:
  - merged **table title cells** (pluralized model name)
  - bold headers
  - sensible column widths
  - multiple tables laid out horizontally on the same sheet with a configurable gap
- **Workbook parsing** into model-specific repositories:
  - `excel_file.cars.all()` → `list[Car]`
  - `excel_file.manufacturing_plants.all()` → `list[ManufacturingPlant]`
- **Validation hooks**
  - column-level `not_null`
  - optional row exclusion rules via `excludes`
  - optional model-level `validate()` method

---

## Installation

### From PyPI (once published)
```bash
pip install excel-orm
````

### From source (uv)

```bash
git clone <your-repo-url>
cd excel-orm
uv sync
```

---

## Quick Start

### 1) Define models using `Column[...]` descriptors

```python
from excel_orm.column import Column, text_column, int_column
from excel_orm.orm import ExcelFile, SheetSpec

class Car:
    make: Column[str] = text_column(header="Make", not_null=True)
    model: Column[str] = text_column(header="Model", not_null=True)
    year: Column[int] = int_column(header="Year", not_null=True)

class ManufacturingPlant:
    name: Column[str] = text_column(header="Factory Name", not_null=True)
    location: Column[str] = text_column(header="Location")
```

### 2) Declare a sheet containing multiple models

Each model becomes its own table on the same worksheet.

```python
sheet = SheetSpec(
    name="Cars",
    models=[Car, ManufacturingPlant],

    # Layout rows
    title_row=1,
    header_row=2,
    data_start_row=3,

    # Horizontal spacing between model tables
    template_table_gap=2,
)
```

### 3) Create an `ExcelFile`, generate a template, then load data

```python
excel_file = ExcelFile(sheets=[sheet])

# Generate a blank template workbook
excel_file.generate_template("car_inventory_template.xlsx")

# Users fill in data in Excel...

# Load the filled workbook into repositories
excel_file.load_data("car_inventory_data.xlsx")

cars = excel_file.cars.all()
plants = excel_file.manufacturing_plants.all()

print(cars[0].make, cars[0].year)
print(plants[0].name, plants[0].location)
```

---

## How It Works

### Repositories

For each model you register, `ExcelFile` creates a repository attribute on the instance using a snake_case pluralized name:

* `Car` → `excel_file.cars`
* `ManufacturingPlant` → `excel_file.manufacturing_plants`

Repositories are simple list-like containers with an `all()` helper:

```python
cars = excel_file.cars.all()  # list[Car]
```

### Multi-table Sheets

A single worksheet can host multiple model tables. During template generation:

* A merged title cell is written above each table (pluralized class name in title case).
* Headers appear under the title.
* Data rows begin at `data_start_row`.
* Tables are placed horizontally with `template_table_gap` blank columns between them.

During parsing:

* The library locates each model table by matching the expected header sequence.
* It reads contiguous rows until a blank row is encountered.

---

## Column Types

### Text

```python
from excel_orm.column import Column, text_column

class Example:
    name: Column[str] = text_column(header="Name", not_null=True, strip=True)
```

* `None` parses to `""` (empty string).
* `strip=True` trims whitespace.

### Integer

```python
from excel_orm.column import Column, int_column

class Example:
    qty: Column[int] = int_column(header="Qty", not_null=True)
```

* `None` or `""` parses to `0`.

### Boolean

```python
from excel_orm.column import Column, bool_column

class Example:
    active: Column[bool] = bool_column(header="Active")
```

Accepted values include:

* True: `true, t, yes, y, 1` (case-insensitive)
* False: `false, f, no, n, 0`
* `None` / empty parses to `False`

Invalid values raise `ValueError`.

### Date

```python
from excel_orm.column import Column, date_column

class Example:
    start_date: Column[date] = date_column(header="Start Date")
```

The date parser supports:

* Excel-native `datetime`/`date` values from `openpyxl`
* ISO strings like `2025-06-01` and `2025-06-01T13:45:00`
* Common business formats including `01-JUN-2025`

Invalid/empty values raise `ValueError`.

---

## Validation

### Column-level: `not_null`

```python
class Car:
    make: Column[str] = text_column(header="Make", not_null=True)
```

If a `not_null=True` column parses to `None` or `""`, a `ValueError` is raised.

### Row exclusion: `excludes`

If you set `excludes`, rows matching those raw values in that column will be skipped.

```python
status: Column[str] = text_column(header="Status")
status.spec.excludes = {"IGNORE", "SKIP"}  # example pattern
```

(If you want a nicer API for excludes, consider adding it directly to the column factory signature.)

### Model-level: `validate()`

If your model defines a `validate(self)` method, it is called after a row is parsed.

```python
class Car:
    make: Column[str] = text_column(header="Make", not_null=True)
    year: Column[int] = int_column(header="Year", not_null=True)

    def validate(self) -> None:
        if self.year < 1886:
            raise ValueError("Invalid car year")
```

---

## Development

### Run tests

```bash
uv run pytest
```

### Lint/format (example)

If you use Ruff:

```bash
uv run ruff check .
uv run ruff format .
```
