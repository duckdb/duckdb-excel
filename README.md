# DuckDB Excel Extension

This extension adds support for the [`TEXT` function](https://support.microsoft.com/en-us/office/text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c) for formatting numbers from Microsoft Excel. 

Example usage:

```sql
SELECT text(1234567.897, '$#.##') AS result;
┌────────────┐
│   result   │
│  varchar   │
├────────────┤
│ $1234567.9 │
└────────────┘
```

# Documentation

See the [Excel page](https://duckdb.org/docs/extensions/excel) in the DuckDB documentation.

# XLSX Files

## Reading XLSX Files

`.xlsx` files can be read using the `read_xlsx` function. The following named parameters are supported.

__Options__:

| Option | Type | Default|  Description |
| --- | --- | --- | --- |
| `header` | `BOOLEAN` | _automatically inferred_  | Whether to treat the first row as containing the names of the resulting columns |
| `sheet`| `VARCHAR` | _automatically inferred_ | The name of the sheet in the xlsx file to read. Default is the first sheet. |
| `all_varchar` | `BOOLEAN` | `false` | Whether to read all cells as containing `VARCHAR`s. |
| `ignore_errors` | `BOOLEAN` | `false` | Whether to ignore errors and silently replace cells that cant be cast to the corresponding inferred column type with `NULL`'s. |
| `range` | `VARCHAR` |  _automatically inferred_ | The range of cells to read. For example, `A1:B2` reads the cells from A1 to B2. If not specified the resulting range will be inferred as rectangular region of cells between the first row of consecutive non-empty cells and the first empty row spanning the same columns |
| `stop_at_empty` | `BOOLEAN` | `false/true` | Whether to stop reading the file when an empty row is encountered. If an explicit `range` option is provided, this is `false` by default, otherwise `true` | 
| `empty_as_varchar` | `BOOLEAN` | `false` | Whether to treat empty cells as `VARCHAR` instead of `DOUBLE` when trying to automatically infer column types |

__Example usage__:

```sql
SELECT * FROM read_xlsx('test.xlsx', header 'true');
----
┌────────┬────────┐
│   a    │   b    │
│ double │ double │
├────────┼────────┤
│    1.0 │    2.0 │
│    3.0 │    4.0 │
└────────┴────────┘

-- Alternatively, we can use a xlsx file as a "replacement scans" and select from it immediately
-- but without being able to pass options.

SELECT * FROM 'test.xlsx';
----
┌────────┬────────┐
│   a    │   b    │
│ double │ double │
├────────┼────────┤
│    1.0 │    2.0 │
│    3.0 │    4.0 │
└────────┴────────┘
```

## Writing XLSX Files

Writing `.xlsx` files is supported using the `COPY` statement with `XLSX` given as the format. The following additional parameters are supported.

__Options__:

| Option | Type | Default   | Description                                                                          |
| --- | --- |-----------|--------------------------------------------------------------------------------------|
| `header` | `BOOLEAN` | `false`   | Whether to write the column names as the first row in the sheet                      |
| `sheet`| `VARCHAR` | `Sheet1`  | The name of the sheet in the xlsx file to write.                                     |
| `sheet_row_limit` | `INTEGER` | `1048576` | The maximum number of rows in a sheet. An error is thrown if this limit is exceeded. |
| `mode` | `VARCHAR` | `create` | How to handle the target file. `create` writes a fresh single-sheet xlsx (overwriting any existing file). `append` adds a new sheet to an existing xlsx (errors if a sheet with the same name already exists). `replace` overwrites a single sheet by name in an existing xlsx (errors if no sheet with that name is found). |

__Example usage__:

```sql
CREATE TABLE test AS SELECT * FROM (VALUES (1, 2), (3, 4)) AS t(a, b);
COPY test TO 'test.xlsx' (format 'xlsx', header 'true');
```

### Writing multiple sheets to one xlsx

Successive `COPY` statements with `mode 'append'` add new sheets to the same workbook. The first call creates the file; subsequent calls add sheets. `mode 'replace'` overwrites a single named sheet, leaving every other sheet untouched.

```sql
COPY (SELECT 1 AS x, 'one'   AS y) TO 'multi.xlsx' (format 'xlsx', header true, sheet 'A');
COPY (SELECT 2 AS x, 'two'   AS y) TO 'multi.xlsx' (format 'xlsx', header true, sheet 'B', mode 'append');
COPY (SELECT 3 AS x, 'three' AS y) TO 'multi.xlsx' (format 'xlsx', header true, sheet 'C', mode 'append');

-- Replace sheet B in place — sheets A and C are unchanged
COPY (SELECT 999 AS x, 'replaced' AS y) TO 'multi.xlsx' (format 'xlsx', header true, sheet 'B', mode 'replace');
```

Notes:
- `mode 'append'` against a path that doesn't exist yet falls back to `create`, so the first call in a multi-`COPY` sequence works the same way regardless of whether you pass `mode 'append'` or omit it.
- `mode 'replace'` keeps the existing sheet's position, internal id, and file path stable, so external references in the workbook (charts, defined names, pivot caches) remain valid after the rewrite.
- Appending into an xlsx authored by Excel, openpyxl, or another tool is supported: the existing workbook structure, styles, shared strings, and all unrelated sheets are preserved verbatim.
- Use `mode 'append'` for adding sheets, `mode 'replace'` for overwriting one. Setting an invalid `mode` value raises an error at bind time.
- `mode 'append'` and `mode 'replace'` only work for **local files**. They need to read the existing workbook and atomically swap a rebuilt copy back into place, which remote filesystems (`s3://`, `https://`, `gcs://`, `azure://`, …) don't support; using either against a remote path raises an error before the destination is touched. Write the workbook to a local path and upload it instead. (`mode 'create'` — the default — works with remote paths as usual, since it just writes a fresh file.)

## Type Conversions and Inference

Because XLSX files only really support storing strings and numbers, the equivalent of `VARCHAR` and `DOUBLE`, the following type conversions are applied when writing XLSX files.
- Numeric types are cast to `DOUBLE` when writing to an XLSX file.
- Temporal types (`TIMESTAMP`, `DATE`, `TIME`, etc.) are converted to excel "serial" numbers, that is the number of days since 1900-01-01 for dates and the fraction of a day for times. These are then styled with a "number format" so that they appear as dates or times in Excel.   
- `TIMESTAMP_TZ` and `TIME_TZ` are cast to UTC `TIMESTAMP` and `TIME` respectively, with the timezone information being lost.
- `BOOLEAN`s are converted to `1` and `0`, with a "number format" applied to make them appear as `TRUE` and `FALSE` in Excel.
- All other types are cast to `VARCHAR` and then written as text cells.

When reading XLSX files, almost everything is read as either `DOUBLE` or `VARCHAR` depending on the Excel cell type. However, there are some caveats.
- We try to infer `TIMESTAMP`, `TIME`, `DATE` and `BOOLEAN` types when possible based on the cell format.
- We infer text cells containing `TRUE` and `FALSE` as `BOOLEAN`, but that is the only type-inference we do that is based on the actual content of the cell.
- Empty cells are considered to be `DOUBLE` by default, unless the `empty_as_varchar` option is set to `true`, in which case they are typed as `VARCHAR`.

If the `all_varchar` option is set to `true`, none of the above applies and all cells are read as `VARCHAR`.

When no types are specified explicitly, (e.g. when using the `read_xlsx` function instead of `COPY TO ... FROM '<file>.xlsx'`) 
the types of the resulting columns are inferred based on the first "data" row in the sheet, that is:
- If no explicit range is given
  - The first row after the header if a header is found or forced by the `header` option
  - The first non-empty row in the sheet if no header is found or forced
- If an explicit range is given
  - The second row of the range if a header is found in the first row or forced by the `header` option
  - The first row of the range if no header is found or forced 

This can sometimes lead to issues if the first "data row" is not representative of the rest of the sheet (e.g. it contains empty cells) in which case the `ignore_errors` or `empty_as_varchar` options can be used to work around this. 
Alternatively, when the `COPY TO ... FROM '<file>.xlsx'` syntax is used, no type inference is done and the types of the resulting columns are determined by the types of the columns in the table being copied to. All cells will simply be converted by casting from `DOUBLE` or `VARCHAR` to the target column type.
