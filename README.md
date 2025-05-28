# Excel Utilities

This repository includes a script for adjusting totals in an Excel sheet by subtracting a specific model from the aggregated values.

## Script: `add_row.py`

`add_row.py` inserts a new row below each occurrence of a total row, subtracting a specified model's row. The new row contains formulas referencing the original total and model rows so that calculations remain dynamic.

### Usage

```bash
python add_row.py <excel_file> <sheet_name> <total_row_name> <exclude_row_name> [--label LABEL]
```

- `excel_file`: Path to the Excel workbook to modify. It will be overwritten.
- `sheet_name`: Name of the sheet containing the data.
- `total_row_name`: Label used in the first column for total rows.
- `exclude_row_name`: Label in the first column of the model row you want to subtract from the total.
- `--label LABEL`: (optional) Custom label for the inserted rows. You may include `{total}` and `{exclude}` placeholders that will be replaced with the provided labels.

The script searches the first column for the specified labels and inserts a new row below each total row, containing formulas that subtract the exclude row from the total row across all columns.

### Example

```bash
python add_row.py data.xlsx "Sales" "Total" "ModelA" --label "S25 except edge"
```

This command will update `data.xlsx`, inserting rows labelled `S25 except edge` beneath each `Total` row in the `Sales` sheet.
