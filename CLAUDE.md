# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Python ETL pipeline that unifies 19 years (2005-2023) of Spanish hunting statistics from MITECO's annual Excel files into clean, analysis-ready CSVs. The project is in Spanish.

The goal is to remove the data-access bottleneck for researchers: MITECO publishes data year-by-year in heterogeneous Excel formats, making longitudinal analysis impractical without significant manual work. This pipeline automates the download, parsing, normalization, and unification into a single coherent dataset.

## Commands

```bash
# Install dependencies
pip install pandas openpyxl xlrd

# Download Excel files from MITECO + process
python unificador.py --todo

# Process already-downloaded files (in raw_data/)
python unificador.py --procesar

# Process specific years (CAUTION: overwrites full CSVs, only useful for debugging)
python unificador.py --procesar --año 2023 2022

# Download only
python unificador.py --descargar

# Run integrity audit (compares CSVs against original Excel files)
python auditoria.py

# Run unit tests (one independent check per year, compares Excel totals vs CSV)
python test_unitarios.py

# Windows UTF-8 mode (if encoding issues)
python -X utf8 unificador.py --procesar
```

## Verification

There are two independent verification mechanisms:

- **`test_unitarios.py`** (~2900 lines): One self-contained `check_XXXX()` function per year (2005-2023). Each reads the original Excel, sums values using its own logic, and compares against the output CSVs. Runs in ~2 minutes. **All 19 years must pass.**

- **`auditoria.py`** (~1640 lines): Four-level audit that re-reads Excel files and compares against generated CSVs:
  - **REPARSE** (184 checks): row-by-row re-parsing comparison.
  - **TOTALES** (109 checks): sum verification against Excel.
  - **SAMPLING** (184 checks): random spot-checks with two seeds.
  - **INTEGRIDAD** (18 checks): null keys, duplicates, completeness.

  Current status: REPARSE 184/184, TOTALES 109/109, INTEGRIDAD 18/18. SAMPLING has 7 known failures due to row-matching limitations in the verifier for edge cases (CCAAs without provincial breakdown, Castilla-La Mancha sueltas 2023). These are **verifier limitations, not data errors** — the data is validated by the other three levels.

## Architecture

### Data Flow

`raw_data/*.xls[x]` → `unificador.py` (parse → normalize → unify) → `output/*.csv`

`auditoria.py` and `test_unitarios.py` verify `output/*.csv` against `raw_data/*.xls[x]`

### unificador.py (~1970 lines)

The core logic is driven by **YEAR_CONFIG** (~lines 280-486), a dictionary mapping each year (2005-2023) to its sheet names, header rows, and which parser function to use. This is necessary because MITECO changed Excel formats significantly across years.

Key sections:
- **Normalization tables and functions** (~lines 109-270): `CCAA_NORM`, `PROV_NORM`, `ESPECIE_NORM`, `CATEGORIA_NORM`, `PROV_TO_CCAA` dictionaries with `norm_*()` functions that canonicalize region/province/species names.
- **Parser functions** (~lines 643-1373): 12+ format-specific parsers (e.g., `parse_licencias_ccaa()`, `parse_capturas()`, `parse_terrenos_paired()`, `parse_capturas_2005()`). Each handles a specific Excel layout variant.
- **Router** (`parse_year()`, ~lines 1530-1587): Selects correct parsers per year based on YEAR_CONFIG.
- **Pipeline** (~lines 1606-1645): Iterates years newest-to-oldest, accumulates DataFrames per table type, concatenates, and writes CSVs.
- **Quality checks** (~lines 1736-1904): Validates canonical values, checks for duplicates and negative values.

### Key parser patterns

- **Forward-fill**: Excel cells use merged/empty cells meaning "same as above" for CCAA, province, and tipo_procedencia. Parsers track state variables (`ccaa_act`, `prov_act`) to fill these.
- **Footnote entries**: Some years (2009-2012) have entries like "Álava (2010)" that are prior-year footnotes, not real data. The `parse_licencias_prov` parser skips these with a `"Name (YYYY)"` pattern check.
- **CCAAs without provincial breakdown**: Some CCAAs (e.g., Aragón 2014) list provinces with empty values. The parser emits a synthetic row with `provincia="(Total CCAA)"` using the CCAA-level total.
- **`_cualificar_otras()`**: Post-processes generic "Otras"/"Otros" species by appending category name (e.g., "Otros Caza Mayor").
- **`_build_col_map`**: Handles double-header Excel format (category row + species row). `sueltas_flat` (2007) has single header and needs its own column-building logic.
- **PROV_TO_CCAA filtering**: Provincial format (2005-2006) lists both provinces AND CCAAs; CCAAs are subtotals that must be filtered using `PROV_TO_CCAA.get(norm_prov(c0))`.

### auditoria.py (~1640 lines)

Independent verification script that re-reads Excel files and compares against generated CSVs at multiple levels of granularity.

### test_unitarios.py (~2900 lines)

Unit tests with one `check_XXXX()` function per year. Shared utilities are minimal: `_sn` (safe numeric), `_csv_año`, `_ok`, `_fail`, `_check`. The `_check_xls_standard()` helper covers 2008-2013 (identical format). Each year's function is otherwise self-contained.

### Key Design Principle

**Conservative normalization**: unify format only, never make analytical decisions. Missing data stays as NaN, unreported regions are simply absent, species taxonomy changes are preserved as-is.

### Output Tables

Six data tables plus audit results:

| Table | Rows | Period | Description |
|-------|-----:|--------|-------------|
| `licencias_ccaa.csv` | 136 | 2016-2023 | Licencias a nivel CCAA (fuente directa) |
| `licencias_prov.csv` | 484 | 2005-2015 | Licencias por provincia |
| `capturas.csv` | 13,418 | 2005-2023 | Capturas por CCAA/prov/categoría/especie |
| `sueltas.csv` | 3,606 | 2006-2023 | Sueltas por tipo procedencia/categoría/especie |
| `produccion.csv` | 1,421 | 2006-2023 | Producción cinegética por categoría/especie |
| `terrenos.csv` | 3,078 | 2006-2023 | Terrenos cinegéticos (nº cotos + superficie) |

All share common columns: `anio`, `ccaa`, `provincia`, `categoria`, `especie`.

## Adding Support for New Years

When MITECO publishes a new year's data:
1. Add the download URL and year config entry to `YEAR_CONFIG` in `unificador.py`
2. Check if the Excel format matches an existing parser or needs a new one
3. Update normalization dictionaries if new region/species name variants appear
4. Add a `check_XXXX()` function to `test_unitarios.py`
5. Run `python unificador.py --procesar` (always without `--año` to regenerate all CSVs)
6. Run `python test_unitarios.py` — all 19+ years must pass
7. Run `python auditoria.py` to verify integrity
