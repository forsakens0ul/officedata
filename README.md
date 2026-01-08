# OfficeMind 📊📑📄

基于 LangGraph 的 **Office 文档智能数据分析助手**，支持 **Excel、PowerPoint、Word** 三种格式的自然语言查询、多轮对话、流式输出、**ECharts 图表可视化**和可视化思考过程。

![Python](https://img.shields.io/badge/Python-3.10+-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![FastAPI](https://img.shields.io/badge/FastAPI-Latest-009688)
![LangChain](https://img.shields.io/badge/LangChain-Latest-orange)

## 🌟 v1.0 全新升级

**ExcelMind 现已升级为 OfficeMind！**

- ✨ **全新支持 PowerPoint (.pptx)** - 提取幻灯片文本、表格和演讲者备注
- ✨ **全新支持 Word (.docx)** - 提取文档段落、标题和表格
- ✨ **统一分析框架** - 三种文档类型使用相同的数据分析工具
- ✨ **智能文档识别** - LLM 自动识别文档类型并选择合适的查询方式

## 🎬 演示视频

[![OfficeMind 演示视频](https://img.shields.io/badge/📺_点击观看-Bilibili视频-00A1D6?style=for-the-badge&logo=bilibili)](https://www.bilibili.com/video/BV1VKvWBgEF1/)

> 视频展示了核心功能：自然语言查询、智能联表、知识库检索等（原 ExcelMind 演示，新版本功能更强大）

## ✨ 功能亮点

### 📂 支持的文档格式

| 格式 | 扩展名 | 功能 |
|------|--------|------|
| **Excel** | `.xlsx`, `.xls`, `.xlsm` | 表格数据分析、统计聚合、图表可视化 |
| **PowerPoint** | `.pptx`, `.ppt` | 幻灯片文本提取、内容搜索、备注分析 |
| **Word** | `.docx`, `.doc` | 文档结构分析、段落搜索、标题提取 |

### 🎯 核心能力

- **自然语言查询**: 用中文直接提问，无需编写代码或公式
  - Excel: "统计各分公司的销售额"
  - PowerPoint: "第 3 张幻灯片讲了什么？"
  - Word: "文档中提到了哪些安装步骤？"
- **多轮对话**: 支持上下文关联的连续追问（如"和上个月相比呢？"）
- **流式输出**: 实时显示 AI 思考过程和回答，响应更流畅
- **智能工具调用**: 自动选择合适的数据分析工具，展示完整推理链路

### 🛠️ 丰富的数据分析工具

所有工具支持 **Excel、PowerPoint、Word** 三种文档格式！

| 工具 | 功能 | 特性 | 适用场景 |
|------|------|------|----------|
| `filter_data` | 筛选+排序 | 支持多条件 AND、排序、指定返回列 | Excel 数据筛选 / PPT 查询特定幻灯片 / Word 查找特定段落 |
| `aggregate_data` | 聚合统计 | 支持先筛选再聚合 | Excel 数值统计 / 统计 PPT 幻灯片数量 |
| `group_and_aggregate` | 分组聚合 | 支持筛选后分组 | Excel 分组统计 / Word 按标题层级统计 |
| `search_data` | 关键词搜索 | 可限制搜索范围 | **PPT/Word 文本搜索的最佳工具** |
| `get_column_stats` | 列统计 | 支持筛选后统计 | 所有文档类型的数据统计 |
| `get_unique_values` | 唯一值 | 支持筛选后获取 | 所有文档类型 |
| `get_data_preview` | 数据预览 | 快速查看数据 | 所有文档类型 |
| `get_current_time` | 获取时间 | 处理相对时间查询 | 时间相关查询 |
| `calculate` | 数学计算 | 批量精确计算 | Excel 数值计算 |
| `generate_chart` | **图表生成** | ECharts 可视化，AI 自动推荐图表类型 | Excel 数据可视化 |

### 📊 文档类型特性

#### Excel 表格分析
- 传统表格数据的完整分析能力
- 数值统计、分组聚合、图表可视化
- 多表联接和协同分析

#### PowerPoint 文本提取
- **幻灯片内容**: 自动提取所有文本内容（标题、正文、列表）
- **演讲者备注**: 提取演讲者备注内容
- **表格数据**: 提取 PPT 中的表格（存储为 JSON）
- **结构化查询**: 按幻灯片编号、内容类型筛选

**PowerPoint 数据结构**:
```
每行 = 一个内容元素
列: slide_number（幻灯片编号）, slide_title（标题）, content_type（类型）,
    content_text（文本）, content_hierarchy（层级）, speaker_notes（备注）
```

**查询示例**:
- "第 3 张幻灯片讲了什么？" → 按 `slide_number` 筛选
- "哪些幻灯片提到了销售？" → 在 `content_text` 中搜索关键词
- "所有幻灯片的标题" → 筛选 `content_type == 'title'`

#### Word 文档分析
- **文档结构**: 识别标题层级（Heading 1-9）
- **段落提取**: 按顺序提取所有段落
- **表格数据**: 提取 Word 中的表格（存储为 JSON）
- **格式信息**: 记录粗体、斜体等格式标记

**Word 数据结构**:
```
每行 = 一个段落或表格
列: paragraph_number（段落号）, content_type（类型）, content_text（文本）,
    hierarchy_level（标题层级）, style_name（样式）, is_bold, is_italic
```

**查询示例**:
- "文档有哪些一级标题？" → 筛选 `hierarchy_level == 1`
- "找到安装相关的段落" → 在 `content_text` 中搜索"安装"
- "列出所有粗体文本" → 筛选 `is_bold == True`

### 🔄 多文档协同

- **多文档管理**: 同时上传和管理多个 Office 文档（Excel + PPT + Word）
- **智能联表**: AI 自动分析 Excel 表结构，通过 `🤖 智能联表` 功能一键生成连接建议
- **灵活连接**: 支持多字段（复合键）连接，以及 Inner/Left/Right/Outer 等多种连接方式
- **上下文感知**: 对话时明确显示当前所在的文档上下文
- **文档类型标识**: 前端显示文档类型图标（📊 Excel, 📊 PPT, 📄 Word）

### 📚 本地知识库

- **私有知识存储**: 存储业务规则、字段说明、操作指南等私有知识
- **向量检索**: 基于 Chroma 向量数据库，使用 Embedding 模型进行语义检索
- **智能召回**: 对话时自动检索相关知识，注入到 Prompt 提升回答质量
- **Web 管理**: 右侧面板可视化管理知识条目，支持增删改查和文件上传
- **持久化存储**: 知识向量化后自动持久化，重启不丢失

### 📈 ECharts 图表可视化

- **多图表类型**: 支持柱状图、折线图、饼图、散点图、雷达图、漏斗图
- **AI 自动推荐**: 根据数据特征智能推荐最合适的图表类型
- **交互式图表**: 基于 ECharts 5.5，支持悬停提示、图例切换、响应式布局
- **自然语言触发**: 直接说"帮我画个图表"或"可视化销售数据"即可生成

### 🦺 现代化 Web 界面

- **双主题模式**: 支持亮色/暗色主题一键切换，自动记忆用户偏好
- **侧边栏管理**: 清晰的文档列表和操作入口
- **拖拽上传**: 支持多文件拖拽上传（Excel + PPT + Word），带进度提示
- **实时预览**: 上传后即时显示数据结构和预览
- **思考可视化**: 展示 AI 的推理过程（Chain of Thought）
- **工具调用展示**: 透明显示每一步工具调用和结果（美化版）
- **Markdown 渲染**: 完美支持表格、代码块等格式

### 🛡️ 安全与稳定

- **意图过滤**: 自动拒绝与文档数据无关的闲聊
- **类型兼容**: 工具参数支持多种数据类型（字符串、数值、日期）
- **模糊匹配**: 日期字段支持前缀匹配（如 "202511" 匹配 "20251104"）
- **高迭代限制**: 支持复杂任务的多步推理（最多 50 次工具调用）

## 🚀 快速开始

### 环境要求

- Python 3.10+
- pip 或 [uv](https://github.com/astral-sh/uv)（推荐）

### 1. 克隆项目

```bash
git clone https://github.com/stark-456/ExcelMind.git
cd ExcelMind
```

### 2. 安装依赖

```bash
# 使用 pip
pip install -e .

# 或使用 uv (推荐)
uv sync
```

**必需依赖**:
- `python-pptx>=0.6.23` - PowerPoint 文件处理
- `python-docx>=1.1.0` - Word 文件处理
- `pandas>=2.0.0` - 数据处理
- `langchain>=0.3.0`, `langgraph>=0.2.0` - LLM 框架
- `fastapi>=0.115.0`, `uvicorn>=0.32.0` - Web 服务

### 3. 配置

复制并编辑 `config.yaml`:

```bash
cp config.example.yaml config.yaml
```

**基础配置**:

```yaml
model:
  active: "default"
  providers:
    default:
      provider: "openai"
      model_name: "gpt-4"  # 推荐使用 gpt-4 或更高版本
      api_key: "${OPENAI_API_KEY}"  # 或直接填写
      base_url: "https://api.openai.com/v1"
      temperature: 0.1
      max_tokens: 4096

# 文档配置（新增）
document:
  max_preview_rows: 10
  pptx_extract_notes: true    # 是否提取 PPT 备注
  pptx_extract_tables: true   # 是否提取 PPT 表格
  docx_extract_tables: true   # 是否提取 Word 表格
  docx_preserve_formatting: true  # 是否保留格式信息

# Embedding 配置（知识库）
embedding:
  active: "default"
  providers:
    default:
      model: "text-embedding-ada-002"
      api_url: "https://api.openai.com/v1"
      api_key: "${OPENAI_API_KEY}"

server:
  host: "0.0.0.0"
  port: 8000
```

**环境变量配置**:

```bash
export OPENAI_API_KEY="your-api-key"
export OPENAI_BASE_URL="your-api-base-url"  # 可选
```

### 4. 启动服务

```bash
# Web 服务模式（推荐）
python -m excel_agent.main serve

# 或使用 uv
uv run python -m excel_agent.main serve

# 命令行模式（仅支持 Excel）
python -m excel_agent.main cli --excel your_file.xlsx
```

### 5. 使用

打开浏览器访问 `http://localhost:8000`:

1. **上传文档**: 拖拽或点击上传 Excel、PowerPoint 或 Word 文件
2. **提出问题**: 在聊天框输入自然语言问题
3. **查看结果**: 查看 AI 的思考过程和分析结果

## 📡 API 接口

启动服务后访问 `http://localhost:8000/docs` 查看完整 Swagger 文档。

### 主要接口

| 接口 | 方法 | 描述 |
|------|------|------|
| `/` | GET | Web 界面 |
| `/upload` | POST | 上传文档（Excel/PPT/Word） |
| `/load` | POST | 通过路径加载文档 |
| `/chat/stream` | POST | 流式对话（推荐） |
| `/chat` | POST | 非流式对话 |
| `/status` | GET | 获取当前状态 |
| `/tables` | GET | 获取所有已加载的文档 |
| `/tables/active` | PUT | 设置活跃文档 |
| `/tables/{table_id}` | DELETE | 删除文档 |
| `/reset` | POST | 重置 Agent |

### 请求示例

```bash
# 上传 PowerPoint 文件
curl -X POST "http://localhost:8000/upload" \
  -F "file=@presentation.pptx"

# 上传 Word 文件
curl -X POST "http://localhost:8000/upload" \
  -F "file=@document.docx"

# 流式对话
curl -X POST "http://localhost:8000/chat/stream" \
  -H "Content-Type: application/json" \
  -d '{
    "message": "第 3 张幻灯片讲了什么？",
    "history": []
  }'
```

## 🏗️ 项目结构

```
OfficeMind/
├── config.yaml                 # 配置文件
├── config.example.yaml         # 配置示例
├── pyproject.toml              # 项目依赖
├── README.md
├── knowledge/                  # 知识库文件目录
│   └── *.md                    # Markdown 格式知识文件
├── .vector_db/                 # Chroma 向量数据库（自动生成）
└── src/
    └── excel_agent/
        ├── __init__.py
        ├── main.py             # 入口
        ├── api.py              # FastAPI 接口
        ├── config.py           # 配置管理
        ├── base_loader.py      # 抽象基类（新增）
        ├── excel_loader.py     # Excel 加载器
        ├── pptx_loader.py      # PowerPoint 加载器（新增）
        ├── docx_loader.py      # Word 加载器（新增）
        ├── document_loader.py  # 多文档管理器（新增）
        ├── graph.py            # LangGraph 工作流
        ├── knowledge_base.py   # 知识库管理
        ├── prompts.py          # 提示词模板
        ├── stream.py           # 流式对话核心
        ├── tools.py            # 数据分析工具
        └── frontend/
            └── index.html      # Web 界面
```

## 📝 使用示例

### Excel 数据分析

```
用户：这个表有多少行数据？
助手：该表共有 15,234 行数据。

用户：按分公司统计移动新增用户总数
助手：[调用 group_and_aggregate 工具]
      各分公司移动新增用户统计如下：
      | 分公司 | 移动新增用户 |
      |--------|-------------|
      | 东城   | 45,678      |
      | 西城   | 38,901      |
      | ...    | ...         |

用户：西城的明细呢？
助手：[理解上下文，调用 filter_data]
      西城分公司的详细数据如下：...

用户：用饼图展示各分公司的占比
助手：[调用 generate_chart 工具]
      📊 已生成饼图，共 8 个数据点。
      [交互式 ECharts 饼图显示]
```

### PowerPoint 内容查询

```
用户：这个 PPT 一共有多少张幻灯片？
助手：[调用 get_unique_values 工具，统计 slide_number]
      这个演示文稿共有 15 张幻灯片。

用户：第 3 张幻灯片讲了什么？
助手：[调用 filter_data，条件：slide_number == 3]
      第 3 张幻灯片的标题是"市场分析"，主要内容包括：
      - 2024 年市场规模达到 500 亿
      - 同比增长 25%
      - 主要增长来自移动端用户

用户：哪些幻灯片提到了"销售"？
助手：[调用 search_data，在 content_text 中搜索"销售"]
      共有 4 张幻灯片提到了"销售"：
      - 第 5 张："销售策略"
      - 第 7 张："销售数据分析"
      - 第 12 张："销售团队介绍"
      - 第 14 张："销售目标"

用户：第 5 张幻灯片的备注是什么？
助手：[调用 filter_data，查看 speaker_notes 列]
      第 5 张幻灯片的演讲者备注为："重点强调线上线下结合的策略，
      举例说明成功案例。"
```

### Word 文档分析

```
用户：这个文档有哪些章节？
助手：[调用 filter_data，条件：hierarchy_level == 1]
      文档包含以下一级标题（章节）：
      1. 引言
      2. 系统架构
      3. 安装指南
      4. 使用说明
      5. 常见问题

用户：安装指南在哪里？详细说说
助手：[调用 search_data，搜索"安装"]
      "安装指南"是第 3 章，包含以下内容：

      **3.1 环境要求**
      - Python 3.10+
      - 操作系统：Windows/Linux/macOS

      **3.2 安装步骤**
      1. 下载安装包
      2. 解压到指定目录
      3. 运行 install.sh

      **3.3 验证安装**
      运行 `python --version` 验证...

用户：文档中有几个表格？
助手：[调用 filter_data，条件：content_type == 'table']
      文档中共有 3 个表格，分别位于：
      - 段落 15：系统配置参数表
      - 段落 28：性能测试结果表
      - 段落 42：API 接口列表表
```

### 多文档协同

```
用户：[上传 sales.xlsx、report.pptx、manual.docx]

用户：报告 PPT 的主题是什么？
助手：[自动识别 report.pptx]
      报告的主题是"2024 年度销售总结与 2025 年规划"...

用户：切换到销售表格，统计各区域销售额
助手：[切换到 sales.xlsx，调用数据分析工具]
      | 区域 | 销售额 |
      |------|--------|
      | 华北 | 1.2亿  |
      | 华东 | 1.5亿  |
      ...

用户：手册里有没有提到培训？
助手：[切换到 manual.docx，搜索"培训"]
      手册中第 6 章"员工培训"详细介绍了培训内容...
```

## 🐳 Docker 部署

```dockerfile
FROM python:3.11-slim

WORKDIR /app

# 安装系统依赖
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# 复制项目文件
COPY . .

# 安装 Python 依赖
RUN pip install --no-cache-dir -e .

EXPOSE 8000

# 启动服务
CMD ["python", "-m", "excel_agent.main", "serve"]
```

**构建和运行**:

```bash
# 构建镜像
docker build -t officemind .

# 运行容器
docker run -p 8000:8000 \
  -e OPENAI_API_KEY=your-api-key \
  -v $(pwd)/knowledge:/app/knowledge \
  -v $(pwd)/.vector_db:/app/.vector_db \
  officemind
```

## 🔧 高级配置

### 多提供商配置

支持同时配置多个 LLM 提供商，快速切换：

```yaml
model:
  active: "openai"  # 当前使用的提供商
  providers:
    openai:
      provider: "openai"
      model_name: "gpt-4"
      api_key: "${OPENAI_API_KEY}"
      description: "OpenAI GPT-4"

    anthropic:
      provider: "anthropic"
      model_name: "claude-3-opus-20240229"
      api_key: "${ANTHROPIC_API_KEY}"
      description: "Anthropic Claude 3 Opus"

    local:
      provider: "openai"  # OpenAI 兼容接口
      model_name: "qwen-plus"
      api_key: "sk-xxx"
      base_url: "http://localhost:8000/v1"
      description: "本地部署模型"
```

### Embedding 模型配置

```yaml
embedding:
  active: "openai"
  providers:
    openai:
      model: "text-embedding-ada-002"
      dims: 1536
      api_url: "https://api.openai.com/v1"
      api_key: "${OPENAI_API_KEY}"

    local:
      model: "bge-large-zh"
      dims: 1024
      api_url: "http://localhost:8001/v1"
      api_key: "empty"
```

## 🧪 开发

```bash
# 安装开发依赖
pip install -e ".[dev]"

# 运行测试
pytest tests/

# 代码格式化
ruff format .

# 类型检查
mypy src/
```

## 📊 技术架构

### 核心技术栈

- **LangGraph**: Agent 工作流编排
- **LangChain**: LLM 应用框架
- **FastAPI**: 高性能 Web 框架
- **pandas**: 数据处理
- **python-pptx**: PowerPoint 文件处理
- **python-docx**: Word 文件处理
- **ChromaDB**: 向量数据库
- **ECharts**: 数据可视化

### 架构设计

```
┌─────────────────┐
│   Web 前端      │
│  (index.html)   │
└────────┬────────┘
         │ HTTP/SSE
┌────────▼────────┐
│   FastAPI       │
│   (api.py)      │
└────────┬────────┘
         │
┌────────▼────────┐
│   LangGraph     │
│   (graph.py)    │
└────┬─────┬──────┘
     │     │
     │     └──────────┐
     │                │
┌────▼─────┐   ┌─────▼──────┐
│  Tools   │   │ Knowledge  │
│(tools.py)│   │   Base     │
└────┬─────┘   └────────────┘
     │
┌────▼──────────────────┐
│   Document Loaders    │
├───────────────────────┤
│ • BaseDocumentLoader  │
│ • ExcelLoader         │
│ • PPTXLoader          │
│ • DOCXLoader          │
│ • MultiDocumentLoader │
└───────────────────────┘
```

### 数据流

```
用户问题 → FastAPI → LangGraph → LLM
                         ↓
                    工具选择
                         ↓
        ┌────────────────┼────────────────┐
        ▼                ▼                ▼
    Excel Tools    PPT Search       Word Search
        │                │                │
        └────────────────┴────────────────┘
                         ↓
                  Loader 获取数据
                         ↓
                   pandas 处理
                         ↓
                   返回结果 → 用户
```

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

### 贡献指南

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

## 📄 License

[MIT License](LICENSE)

## 🗺️ Roadmap

- [ ] PDF 文件支持
- [ ] 图片内容分析（基于视觉模型）
- [ ] 跨文档联合查询
- [ ] 文档对比分析
- [ ] 导出分析报告
- [ ] 多语言支持

## 💬 社区交流

![QQ群名片](docs/card.png)

---

**Made with ❤️ using LangGraph, FastAPI, and LLM**

**⭐ 如果这个项目对你有帮助，请给我们一个 Star！**
