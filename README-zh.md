<div align="center">

# SKILLful MCP Docx Processor

[English](README.md) | [中文](README-zh.md)

基于 [FastMCP](https://github.com/modelcontextprotocol/python-sdk) 和 [python-docx](https://python-docx.readthedocs.io/) 的 Word 文档处理服务，通过单一网关工具 (`docx_process`) 路由到 **63 个操作**，覆盖文档创建、编辑、格式化、查询的完整生命周期。

专为 [Claude Code](https://docs.anthropic.com/en/docs/claude-code) 设计，通过 MCP 协议和 Claude 的 Skill 系统实现无缝集成。

</div>

---

## 功能概览

| 分类     | 数量     | 说明                      |
| ------ | ------ | ----------------------- |
| 文档管理   | 7      | 创建、打开、保存、另存为、副本、关闭、重新加载 |
| 内容添加   | 8      | 段落、标题、表格、列表、图片、分页符、节、目录 |
| 内容编辑   | 9      | 搜索、替换、预览、章节替换、关键词编辑、删除  |
| 表格操作   | 8      | 行列操作、合并拆分单元格、边框、背景色     |
| 格式设置   | 20     | 页面布局、段落、文本、样式、超链接       |
| 样式管理   | 3      | 创建、修改、列出样式              |
| 批注与引用  | 8      | 书签、批注、脚注、尾注、查询          |
| **合计** | **63** |                         |

---

## 环境要求

- Python >= 3.10
- [uv](https://docs.astral.sh/uv/) 包管理器

## 安装

```bash
# 克隆仓库
git clone https://github.com/Zeke-777/SKILLful-mcp-docx-processor.git
cd SKILLful-mcp-docx-processor

# 安装依赖
uv sync
```

## 配置 MCP 服务器

### Claude Code

适用于Claude Code CLI安装：

```bash
# 项目级（在当前目录创建 .mcp.json）
claude mcp add -s project docx-processor -- uv run --directory /你的项目路径/SKILLful-mcp-docx-processor python server.py

# 或全局级（所有项目可用）
claude mcp add -s user docx-processor -- uv run --directory /你的项目路径/SKILLful-mcp-docx-processor python server.py
```

将 `/你的项目路径/SKILLful-mcp-docx-processor` 替换为你的实际项目路径。

### 手动 JSON 配置

适用于其他 MCP 客户端，或手动创建 `.mcp.json` / `mcp_config.json`：

```json
{
  "mcpServers": {
    "docx-processor": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "/你的项目路径/SKILLful-mcp-docx-processor",
        "python",
        "server.py"
      ],
      "env": {}
    }
  }
}
```

## 安装 Skill

Skill 文件支持项目级和全局两种安装方式，按需选择其一。

### 项目级安装

仅对当前项目生效，Skill 文件位于项目目录内：

```bash
# 在你的目标项目根目录下执行
mkdir -p .claude/skills

# 将 Skill 文件复制到目标项目
cp -r /你的项目路径/SKILLful-mcp-docx-processor/skills/docx-process .claude/skills/
```

### 全局安装

对所有项目生效，Skill 文件位于用户主目录下：

```bash
# 将 Skill 文件复制到全局目录
mkdir -p ~/.claude/skills
cp -r /你的项目路径/SKILLful-mcp-docx-processor/skills/docx-process ~/.claude/skills/
```

安装后目录结构如下：

```
.claude/              # 项目级 或 ~/.claude/  # 全局
└── skills/
    └── docx-process/
        ├── SKILL.md                  # 路由表
        └── references/
            ├── doc-management.md     # 文档管理
            ├── content-operations.md # 内容操作
            ├── table-operations.md   # 表格操作
            ├── formatting.md         # 格式设置
            └── annotations-query.md  # 批注与查询
```

Claude Code 会在调用 `docx_process` 前自动加载 `SKILL.md`，确保每次调用都有完整的路由和参数参考。

---

## 使用方法

配置完成后，`docx_process` 工具即可在 Claude Code 中使用。工具采用单一网关模式：

```
docx_process(route="create_document", params={"file_path": "output.docx"})
docx_process(route="add_heading", params={"text": "你好世界", "level": 1})
docx_process(route="save_document", params={})
```

项目内置了 `docx-process` Skill，每次调用前会自动加载路由表和参数规范，确保调用正确。

## 架构设计

```
server.py          -- MCP 服务器，包含所有 handler 逻辑（2,247 行）
SKILL.md           -- 路由表（每次调用加载）
references/*.md    -- 5 个分类参考文档（按需加载）
```

服务器暴露单一工具 (`docx_process`)，通过 `ROUTE_HANDLERS` 字典分发到对应的处理函数。

## 上下文效率

63 个路由如果全部写入工具的 docstring 或一次性加载，会占用大量上下文窗口。本项目通过 **Skill 按需加载** 机制显著降低上下文开销：

| 加载阶段        | 内容               | 行数          |
| ----------- | ---------------- | ----------- |
| tool schema | docstring（仅参数说明） | 6           |
| 每次调用        | SKILL.md 路由表     | 116         |
| 按需加载        | 单个 reference 文件  | 47-177      |
| **单次总计**    |                  | **169-299** |

对比朴素方案（docstring 含完整路由列表 ~30 行 + 全量 SKILL.md 595 行 = ~625 行），**上下文占用减少 52-73%**。

这意味着：

- **更多空间留给实际任务** — 省下的上下文可用于处理文档内容和业务逻辑
- **响应更快** — 更少的 token 意味着更快的推理速度
- **成本更低** — 每次调用消耗的 token 数大幅减少

---

## 健壮性

- **文件锁检测** — 所有文件操作均检测锁文件
- **参数校验** — 所有路由均校验参数（层级、行列数、颜色、间距等）
- **状态管理** — 覆盖提醒、关闭时清理
- **统一错误处理** — 清晰可操作的错误消息
- **保存重试** — 3 次重试，2 秒间隔

## 已知局限

- 单文档模型（打开新文档会提醒关闭当前文档）
- 不支持段落内富文本（仅支持整段格式化）
- 不支持文本框
- 无并发保护（面向单用户 MCP 场景）

## 致谢

本项目基于 [MCP-Doc](https://github.com/MeterLong/MCP-Doc) 扩展开发，保留了其核心的 `DocxProcessor` 类架构和 21 个基础文档操作，并在此基础上扩展至 63 个路由，新增了文件锁检测、Skill 按需加载等机制。感谢原作者的优秀工作。

## 许可证

MIT
