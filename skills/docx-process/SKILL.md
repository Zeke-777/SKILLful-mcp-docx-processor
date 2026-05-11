---
name: docx-process
description: >
  This skill should be used when the user asks to create, edit, format,
  query, or manipulate Word documents (.docx files). It provides routing
  guidance for the docx_process MCP tool, including all available routes
  and their parameter specifications.
---

# Word Document Processing

Use `docx_process(route="<route>", params={...})` to perform operations.

For detailed parameter specs, read the reference file in the same directory as this skill:
- `references/doc-management.md` — Document management (create, open, save, close, reload)
- `references/content-operations.md` — Content addition and editing (paragraphs, headings, tables, search/replace)
- `references/table-operations.md` — Table operations (rows, columns, merge, split, borders, shading)
- `references/formatting.md` — Formatting (page layout, paragraph, text, styles)
- `references/annotations-query.md` — Annotations (bookmarks, comments, footnotes) and queries

## Route Quick Reference

### Document Management
| Route | Description |
|---|---|
| create_document | Create new document (`file_path`) |
| open_document | Open existing document (`file_path`) |
| save_document | Save to original path |
| save_as_document | Save as new file (`new_file_path`) |
| create_document_copy | Copy with suffix (`suffix`, default `"-副本"`) |
| close_document | Close without saving |
| reload_document | Reload from disk |

### Content Addition
| Route | Description |
|---|---|
| add_paragraph | Add paragraph (`text`, formatting opts) |
| add_heading | Add heading (`text`, `level` 1-9) |
| add_table | Add table (`rows`, `cols`, `data`) |
| add_list_item | Add list item (`text`, `list_type`, `level`) |
| add_image | Insert image (`image_path`, `width`, `height`) |
| add_page_break | Insert page break |
| add_section | Add section (`orientation`, dimensions, margins) |
| add_table_of_contents | Insert TOC field (`title`) |

### Content Editing
| Route | Description |
|---|---|
| search_text | Search text (`keyword`) |
| search_and_replace | Search & replace with preview (`keyword`, `replace_with`, `preview_only`) |
| find_and_replace | Simple find & replace (`find_text`, `replace_text`) |
| replace_section | Replace under heading, keep format (`section_title`, `new_content`, `preserve_title`) |
| edit_section_by_keyword | Replace around keyword, keep format (`keyword`, `new_content`, `section_range`) |
| delete_paragraph | Delete paragraph (`paragraph_index`) |
| delete_text | Delete text range (`paragraph_index`, `start_pos`, `end_pos`) |
| edit_table_cell | Edit cell (`table_index`, `row_index`, `col_index`, `text`) |

### Table Operations
| Route | Description |
|---|---|
| add_table_row | Add row (`table_index`, `data`) |
| add_table_column | Add column (`table_index`, `col_index`, `data`) |
| delete_table_row | Delete row (`table_index`, `row_index`) |
| delete_table_column | Delete column (`table_index`, `col_index`) |
| merge_table_cells | Merge cells (`table_index`, `start_row`, `start_col`, `end_row`, `end_col`) |
| split_table | Split after row (`table_index`, `row_index`) |
| set_table_borders | Set table borders (`table_index`, `color`, `size`) |
| set_cell_shading | Set cell background (`table_index`, `row_index`, `col_index`, `color`) |

### Formatting — Page Layout
| Route | Description |
|---|---|
| set_page_margins | Set margins (`top`, `bottom`, `left`, `right`, `section_index`) |
| set_page_orientation | Set orientation (`orientation`, `section_index`) |
| set_page_size | Set page size (`preset` or `width`+`height`, `section_index`) |
| set_header | Set header text (`text`, `section_index`) |
| set_footer | Set footer text (`text`, `section_index`) |
| set_columns | Set column layout (`num_columns`, `space`, `section_index`) |
| add_page_number | Add PAGE field to footer (`section_index`) |

### Formatting — Paragraph & Text
| Route | Description |
|---|---|
| set_line_spacing | Set line spacing (`spacing`, `paragraph_index`) |
| set_paragraph_spacing | Set spacing & indent (`paragraph_index`, `space_before`, `space_after`, `first_line_indent`) |
| set_paragraph_format | Modify formatting (`paragraph_index`, bold/italic/font/size/color/alignment) |
| set_paragraph_text | Replace text keeping format (`paragraph_index`, `text`) |
| set_paragraph_border | Set borders (`paragraph_index`, `top_style`/`left_style`/`bottom_style`/`right_style` + opts) |
| set_paragraph_shading | Set background (`paragraph_index`, `color`) |
| set_text_highlight | Set highlight (`paragraph_index`, `color`) |
| set_text_strikethrough | Set strikethrough (`paragraph_index`, `double`) |
| set_superscript | Set superscript (`paragraph_index`) |
| set_subscript | Set subscript (`paragraph_index`) |
| set_text_direction | Set direction (`paragraph_index`, `direction`) |
| add_tab_stop | Add tab stop (`paragraph_index`, `position`, `alignment`) |
| add_hyperlink | Add hyperlink (`paragraph_index`, `text`, `url`) |
| set_cell_format | Format cell (`table_index`, `row_index`, `col_index`, formatting opts) |

### Styles
| Route | Description |
|---|---|
| create_style | Create style (`style_name`, `style_type`, formatting opts) |
| modify_style | Modify style (`style_name`, formatting opts) |
| list_styles | List styles (`style_type`) |

### Annotations & Query
| Route | Description |
|---|---|
| add_bookmark | Add bookmark (`paragraph_index`, `bookmark_name`) |
| add_comment | Add comment (`paragraph_index`, `text`, `author`) |
| add_footnote | Add footnote (`paragraph_index`, `text`) |
| add_endnote | Add endnote (`paragraph_index`, `text`) |
| get_document_info | Get metadata |
| get_paragraph | Get paragraph detail (`paragraph_index`) |
| get_table_cell | Get cell content (`table_index`, `row_index`, `col_index`) |
| get_page_info | Get page layout (`section_index`) |
