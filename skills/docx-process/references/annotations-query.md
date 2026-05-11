# Annotations & Query — Parameter Details

## Annotations

### add_bookmark
Add a bookmark to a paragraph.
```json
{"paragraph_index": 0, "bookmark_name": "my_bookmark"}
```
- `paragraph_index` (required), `bookmark_name` (required, must be unique in document)

### add_comment
Add a comment to a paragraph.
```json
{"paragraph_index": 0, "text": "This needs review", "author": "Reviewer"}
```
- `paragraph_index` (required), `text` (required), `author` (optional, default: `"MCP Server"`)

### add_footnote
Add a footnote to a paragraph.
```json
{"paragraph_index": 0, "text": "Footnote content here"}
```
- `paragraph_index` (required), `text` (required)

### add_endnote
Add an endnote to a paragraph.
```json
{"paragraph_index": 0, "text": "Endnote content here"}
```
- `paragraph_index` (required), `text` (required)

## Query

### get_document_info
Get document metadata (paragraph count, table count, styles, etc.).
```json
{}
```

### get_paragraph
Get paragraph text and formatting by index.
```json
{"paragraph_index": 0}
```
- `paragraph_index` (required)
- Returns JSON: `text`, `style`, `alignment`, `runs` (array of `{text, bold, italic, underline, font_size, font_name, color}`)

### get_table_cell
Get table cell content.
```json
{"table_index": 0, "row_index": 0, "col_index": 0}
```
- `table_index`, `row_index`, `col_index` (all required)
- Returns JSON: `text`

### get_page_info
Get page layout details for a section.
```json
{"section_index": 0}
```
- `section_index` (optional, default: `0`)
- Returns JSON: `orientation`, `page_width`, `page_height`, `top_margin`, `bottom_margin`, `left_margin`, `right_margin` (all in cm)
