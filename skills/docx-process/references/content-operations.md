# Content Operations — Parameter Details

## add_paragraph
Add a paragraph with optional formatting.
```json
{
  "text": "Hello World",
  "bold": false, "italic": false, "underline": false,
  "font_size": 12, "font_name": "Arial", "color": "#FF0000",
  "alignment": "left"
}
```
- `text` (required)
- `bold`, `italic`, `underline` (optional, default: `false`)
- `font_size` (optional, in points)
- `font_name` (optional)
- `color` (optional, format: `#RRGGBB`)
- `alignment` (optional): `"left"`, `"center"`, `"right"`, `"justify"`

## add_heading
Add a heading.
```json
{"text": "Chapter 1", "level": 1}
```
- `text` (required)
- `level` (required): 1-9

## add_table
Add a table with optional data.
```json
{"rows": 3, "cols": 2, "data": [["A", "B"], ["C", "D"], ["E", "F"]]}
```
- `rows` (required)
- `cols` (required)
- `data` (optional): 2D array of strings

## add_list_item
Add a bulleted or numbered list item.
```json
{"text": "Item one", "list_type": "bullet", "level": 0}
```
- `text` (required)
- `list_type` (optional): `"bullet"` (default) or `"number"`
- `level` (optional): 0-8, default 0

## add_image
Insert an image.
```json
{"image_path": "path/to/image.png", "width": 10.0, "height": 8.0}
```
- `image_path` (required)
- `width`, `height` (optional, in cm)

## add_page_break
```json
{}
```

## add_section
Add a new section with layout settings. All parameters optional.
```json
{
  "orientation": "landscape",
  "page_width": 29.7, "page_height": 21.0,
  "top_margin": 2.0, "bottom_margin": 2.0, "left_margin": 2.0, "right_margin": 2.0
}
```

## add_table_of_contents
Insert a table of contents field.
```json
{"title": "目录"}
```
- `title` (optional, default: `"目录"`)

## search_text
Search for text in the document (paragraphs and table cells).
```json
{"keyword": "search term"}
```
- `keyword` (required)

## search_and_replace
Search and replace with preview option.
```json
{"keyword": "old", "replace_with": "new", "preview_only": false}
```
- `keyword` (required), `replace_with` (required), `preview_only` (optional, default: `false`)

## find_and_replace
Simple find and replace.
```json
{"find_text": "old", "replace_text": "new"}
```
- `find_text` (required), `replace_text` (required)

## replace_section
Replace content under a section heading, preserving formatting.
```json
{"section_title": "Introduction", "new_content": ["New paragraph 1", "New paragraph 2"], "preserve_title": true}
```
- `section_title` (required)
- `new_content` (required): list of paragraph strings
- `preserve_title` (optional, default: `true`)

## edit_section_by_keyword
Replace paragraphs around a keyword, preserving formatting.
```json
{"keyword": "target", "new_content": ["Replacement text"], "section_range": 3}
```
- `keyword` (required), `new_content` (required), `section_range` (optional, default: `3`)

## delete_paragraph
```json
{"paragraph_index": 5}
```
- `paragraph_index` (required)

## delete_text
Delete text within a paragraph by position range.
```json
{"paragraph_index": 0, "start_pos": 5, "end_pos": 10}
```
- `paragraph_index` (required), `start_pos` (required, 0-based), `end_pos` (required, exclusive)

## edit_table_cell
Edit a table cell's content.
```json
{"table_index": 0, "row_index": 0, "col_index": 0, "text": "New cell text"}
```
- `table_index`, `row_index`, `col_index`, `text` (all required)
