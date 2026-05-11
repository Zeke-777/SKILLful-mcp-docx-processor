# Formatting — Parameter Details

## Page Layout

### set_page_margins
Set page margins in cm.
```json
{"top": 2.54, "bottom": 2.54, "left": 3.18, "right": 3.18, "section_index": 0}
```
- All margin params optional, `section_index` (optional, default: `0`)

### set_page_orientation
```json
{"orientation": "portrait", "section_index": 0}
```
- `orientation` (required): `"portrait"` or `"landscape"`
- `section_index` (optional, default: `0`)

### set_page_size
```json
{"preset": "A4", "section_index": 0}
```
- `preset` (optional): `"A4"`, `"A3"`, `"Letter"`, `"Legal"`
- Or use `width` and `height` (in cm) instead of preset
- `section_index` (optional, default: `0`)

### set_header / set_footer
```json
{"text": "Header Text", "section_index": 0}
```
- `text` (required), `section_index` (optional, default: `0`)

### set_columns
Set column layout for a section.
```json
{"num_columns": 2, "space": 0.5, "section_index": 0}
```
- `num_columns` (required), `space` (optional, column gap in cm), `section_index` (optional, default: `0`)

### add_page_number
Add a PAGE field to a section footer.
```json
{"section_index": 0}
```
- `section_index` (optional, default: `0`)

## Paragraph Formatting

### set_line_spacing
```json
{"spacing": "1.5", "paragraph_index": 0}
```
- `spacing` (required): `"1.0"`, `"1.5"`, `"2.0"`, `"multiple:3"`, `"exact:20"`
- `paragraph_index` (optional): omit to set all paragraphs

### set_paragraph_spacing
```json
{"paragraph_index": 0, "space_before": 6, "space_after": 6, "first_line_indent": 0.74}
```
- `paragraph_index` (required)
- `space_before` (optional, pt), `space_after` (optional, pt), `first_line_indent` (optional, cm)
- At least one of the three spacing params is required

### set_paragraph_format
Modify formatting for existing paragraph. Only passed attributes are changed.
```json
{
  "paragraph_index": 0,
  "bold": true, "italic": false, "underline": false,
  "font_size": 14, "font_name": "Arial", "color": "#FF0000", "alignment": "center"
}
```
- `paragraph_index` (required)
- `bold`, `italic`, `underline` (optional), `font_size` (optional, pt), `font_name` (optional)
- `color` (optional, `#RRGGBB`), `alignment` (optional: `left`/`center`/`right`/`justify`)

### set_paragraph_text
Replace paragraph text, preserving original formatting.
```json
{"paragraph_index": 0, "text": "New text"}
```
- `paragraph_index`, `text` (required)

### set_paragraph_border
```json
{
  "paragraph_index": 0,
  "top_style": "single", "top_size": 4, "top_color": "000000", "top_space": 1,
  "bottom_style": "single"
}
```
- `paragraph_index` (required)
- At least one side required: `top_style`, `left_style`, `bottom_style`, `right_style`
- Border styles: `"single"`, `"double"`, `"dashed"`, `"dotted"`, `"thick"`, etc.
- `*_size` (optional, default: `4`), `*_color` (optional, default: `"000000"`), `*_space` (optional, default: `1`)

### set_paragraph_shading
```json
{"paragraph_index": 0, "color": "FFFF00"}
```
- `paragraph_index` (required), `color` (required, hex without `#`)

## Text Styling

### set_text_highlight
```json
{"paragraph_index": 0, "color": "YELLOW"}
```
- `paragraph_index` (required)
- `color` (required): `YELLOW`, `GREEN`, `CYAN`, `MAGENTA`, `BLUE`, `RED`, `DARK_BLUE`, `DARK_CYAN`, `DARK_GREEN`, `DARK_MAGENTA`, `DARK_RED`, `DARK_YELLOW`, `LIGHT_GRAY`, `DARK_GRAY`, `BLACK`, `WHITE`

### set_text_strikethrough
```json
{"paragraph_index": 0, "double": false}
```
- `paragraph_index` (required), `double` (optional, default: `false`)

### set_superscript / set_subscript
```json
{"paragraph_index": 0}
```
- `paragraph_index` (required)

### set_text_direction
```json
{"paragraph_index": 0, "direction": "ltr"}
```
- `paragraph_index` (required), `direction` (required: `"ltr"` or `"rtl"`)

### add_tab_stop
```json
{"paragraph_index": 0, "position": 5.0, "alignment": "left"}
```
- `paragraph_index` (required), `position` (required, in cm)
- `alignment` (optional): `"left"`, `"center"`, `"right"`, `"decimal"`

### add_hyperlink
```json
{"paragraph_index": 0, "text": "Click here", "url": "https://example.com"}
```
- `paragraph_index`, `text`, `url` (all required)

### set_cell_format
Set formatting for a table cell.
```json
{
  "table_index": 0, "row_index": 0, "col_index": 0,
  "bold": true, "font_size": 12, "font_name": "Arial", "color": "#0000FF"
}
```
- `table_index`, `row_index`, `col_index` (required)
- `bold`, `italic`, `underline`, `font_size`, `font_name`, `color` (all optional)

## Style Management

### create_style
```json
{
  "style_name": "MyHeading", "style_type": "paragraph",
  "bold": true, "font_size": 16, "font_name": "Arial", "color": "#333333"
}
```
- `style_name` (required, must be unique)
- `style_type` (optional, default: `"paragraph"`): `"paragraph"`, `"character"`, `"table"`, `"list"`
- `bold`, `italic`, `font_size`, `font_name`, `color` (all optional)

### modify_style
```json
{"style_name": "MyHeading", "bold": true, "font_size": 14, "color": "#FF0000"}
```
- `style_name` (required), plus any attributes to change

### list_styles
```json
{"style_type": "paragraph"}
```
- `style_type` (optional): `"paragraph"`, `"character"`, `"table"`, `"list"` — omit to list all
