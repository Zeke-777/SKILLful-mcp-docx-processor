<div align="center">

# SKILLful MCP Docx Processor

[English](README.md) | [中文](README-zh.md)

A Word document processing service based on [FastMCP](https://github.com/modelcontextprotocol/python-sdk) and [python-docx](https://python-docx.readthedocs.io/), exposing a single gateway tool (`docx_process`) that routes to **63 operations**, covering the full document lifecycle: creation, editing, formatting, and querying.

Designed for seamless integration with [Claude Code](https://docs.anthropic.com/en/docs/claude-code) via the MCP protocol and Claude's skill system.

</div>

---

## Features

| Category                 | Count  | Description                                                             |
| ------------------------ | ------ | ----------------------------------------------------------------------- |
| Document Management      | 7      | Create, open, save, save-as, copy, close, reload                        |
| Content Addition         | 8      | Paragraphs, headings, tables, lists, images, page breaks, sections, TOC |
| Content Editing          | 9      | Search, replace, preview, section replace, keyword edit, delete         |
| Table Operations         | 8      | Rows, columns, merge/split cells, borders, shading                      |
| Formatting               | 20     | Page layout, paragraph, text, styles, hyperlinks                        |
| Style Management         | 3      | Create, modify, list styles                                             |
| Annotations & References | 8      | Bookmarks, comments, footnotes, endnotes, queries                       |
| **Total**                | **63** |                                                                         |

---

## Requirements

- Python >= 3.10
- [uv](https://docs.astral.sh/uv/) package manager

## Installation

```bash
# Clone the repository
git clone https://github.com/Zeke-777/SKILLful-mcp-docx-processor.git
cd SKILLful-mcp-docx-processor

# Install dependencies
uv sync
```

## Configure MCP Server

### Claude Code

For Claude Code CLI:

```bash
# Project-level (creates .mcp.json in current directory)
claude mcp add -s project docx-processor -- uv run --directory /path/to/SKILLful-mcp-docx-processor python server.py

# Or global (available in all projects)
claude mcp add -s user docx-processor -- uv run --directory /path/to/SKILLful-mcp-docx-processor python server.py
```

Replace `/path/to/SKILLful-mcp-docx-processor` with your actual project path.

### Manual JSON Configuration

For other MCP clients, or to manually create `.mcp.json` / `mcp_config.json`:

```json
{
  "mcpServers": {
    "docx-processor": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "/path/to/SKILLful-mcp-docx-processor",
        "python",
        "server.py"
      ],
      "env": {}
    }
  }
}
```

## Install Skill

Skill files support both project-level and global installation. Choose one as needed.

### Project-level Installation

Takes effect for the current project only. Skill files reside within the project directory:

```bash
# Run in your target project's root directory
mkdir -p .claude/skills

# Copy Skill files to the target project
cp -r /path/to/SKILLful-mcp-docx-processor/skills/docx-process .claude/skills/
```

### Global Installation

Takes effect for all projects. Skill files reside in the user's home directory:

```bash
# Copy Skill files to the global directory
mkdir -p ~/.claude/skills
cp -r /path/to/SKILLful-mcp-docx-processor/skills/docx-process ~/.claude/skills/
```

After installation, the directory structure is as follows:

```
.claude/              # Project-level  or  ~/.claude/  # Global
└── skills/
    └── docx-process/
        ├── SKILL.md                  # Route table
        └── references/
            ├── doc-management.md     # Document management
            ├── content-operations.md # Content operations
            ├── table-operations.md   # Table operations
            ├── formatting.md         # Formatting
            └── annotations-query.md  # Annotations & queries
```

Claude Code automatically loads `SKILL.md` before calling `docx_process`, ensuring every call has the complete route and parameter reference.

---

## Usage

Once configured, the `docx_process` tool becomes available in Claude Code. The tool uses a single gateway pattern:

```
docx_process(route="create_document", params={"file_path": "output.docx"})
docx_process(route="add_heading", params={"text": "Hello World", "level": 1})
docx_process(route="save_document", params={})
```

The built-in `docx-process` skill automatically loads the route table and parameter specifications before each call, ensuring correct usage.

## Architecture

```
server.py          -- MCP server with all handler logic (2,247 lines)
SKILL.md           -- Route table (loaded on every call)
references/*.md    -- 5 category reference docs (loaded on demand)
```

The server exposes one tool (`docx_process`) that dispatches to handler functions via a `ROUTE_HANDLERS` dictionary.

## Context Efficiency

If all 63 routes were written into the tool's docstring or loaded at once, it would consume a large portion of the context window. This project uses **on-demand Skill loading** to significantly reduce context overhead:

| Loading Stage      | Content                          | Lines       |
| ------------------ | -------------------------------- | ----------- |
| Tool schema        | Docstring (parameter specs only) | 6           |
| Every call         | SKILL.md route table             | 116         |
| On demand          | Single reference file            | 47-177      |
| **Total per call** |                                  | **169-299** |

Compared to a naive approach (docstring with full route list ~30 lines + full SKILL.md 595 lines = ~625 lines), **context usage is reduced by 52-73%**.

This means:

- **More room for actual tasks** — saved context can be used for document content and business logic
- **Faster responses** — fewer tokens means faster inference
- **Lower costs** — significantly fewer tokens consumed per call

---

## Robustness

- **File lock detection** for all file operations
- **Parameter validation** across all routes (levels, rows/cols, colors, spacing, etc.)
- **State management** with overwrite warnings and cleanup on close
- **Unified error handling** with clear, actionable messages
- **Save retry** (3 attempts, 2s intervals) for locked files

## Known Limitations

- Single-document model (opening a new document warns about closing the current one)
- No per-run rich text within a paragraph (whole-paragraph formatting only)
- No text box support
- No concurrent multi-user protection (designed for single-user MCP scenarios)

## Acknowledgments

This project is built upon an extension of [MCP-Doc](https://github.com/MeterLong/MCP-Doc), retaining its core `DocxProcessor` class architecture and 21 foundational document operations. It has been expanded to 63 routes with additional features including file lock detection and on-demand Skill loading. Thanks to the original author for the excellent work.

## License

MIT
