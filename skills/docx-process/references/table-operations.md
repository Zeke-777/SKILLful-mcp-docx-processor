# Table Operations — Parameter Details

## add_table_row
Add a row to a table.
```json
{"table_index": 0, "data": ["Cell 1", "Cell 2"]}
```
- `table_index` (required), `data` (optional, string list)

## add_table_column
Add a column to a table.
```json
{"table_index": 0, "col_index": -1, "data": ["Row 1", "Row 2"]}
```
- `table_index` (required), `col_index` (required, use `-1` to append at end), `data` (optional)

## delete_table_row
```json
{"table_index": 0, "row_index": 2}
```
- `table_index`, `row_index` (required)

## delete_table_column
```json
{"table_index": 0, "col_index": 1}
```
- `table_index`, `col_index` (required)

## merge_table_cells
Merge a range of table cells.
```json
{"table_index": 0, "start_row": 0, "start_col": 0, "end_row": 1, "end_col": 1}
```
- `table_index`, `start_row`, `start_col`, `end_row`, `end_col` (all required)

## split_table
Split a table after a specified row.
```json
{"table_index": 0, "row_index": 2}
```
- `table_index`, `row_index` (required)

## set_table_borders
Set borders for an entire table (top, left, bottom, right, insideH, insideV).
```json
{"table_index": 0, "color": "000000", "size": 4}
```
- `table_index` (required)
- `color` (optional, default: `"000000"`, hex without `#`)
- `size` (optional, default: `4`, in eighths of a point)

## set_cell_shading
Set background shading for a table cell.
```json
{"table_index": 0, "row_index": 0, "col_index": 0, "color": "FFFF00"}
```
- `table_index`, `row_index`, `col_index`, `color` (all required, color is hex without `#`)
