# Claude Code 配置与技能库

本仓库存储 Claude Code 的完整配置、技能定义和项目模板，便于复现工作成果和迁移到其他设备。

## 仓库地址
https://github.com/Mrdtmm/PPT-

## 文件结构

```
├── skills/                        # 自定义技能
│   ├── docx-to-ppt-with-images.md  # DOCX图片提取与PPT制作技能
│   └── fix-thesis-citations.md     # 论文引用格式修复技能
├── mcp/                          # MCP配置模板
│   └── README.md                  # MCP服务器配置说明
├── claude.md.example              # CLAUDE.md示例（PolarNet项目）
├── claude.md.template             # CLAUDE.md模板
├── settings.template.json          # 配置模板（请替换Token）
└── README.md
```

## 安装与配置

### 1. 克隆仓库
```bash
git clone https://github.com/Mrdtmm/PPT-.git
```

### 2. 配置 Settings
复制 settings.template.json 到 ~/.claude/settings.json，填入你的 API Token

### 3. 安装技能
复制 skills/ 目录内容到 ~/.claude/skills/

### 4. 安装依赖
```bash
pip install python-docx pillow python-pptx
```

## 技能说明

### docx-to-ppt-with-images
从 Word 文档(.docx)中提取图片素材并生成PPT。适用于学术论文结题汇报。

### fix-thesis-citations
修复Word论文文档中的引用格式问题，包括上标引用添加、双括号清理等。

## MCP Servers

Claude Code 支持通过 MCP 连接外部工具。详细配置见 mcp/README.md

| Server | 用途 |
|--------|------|
| github | GitHub API 操作 |
| gitlab | GitLab API 操作 |
| asana | 项目管理 |
| linear | 问题跟踪 |
| playwright | 浏览器自动化 |

## CLAUDE.md 模板

- claude.md.template - 通用模板
- claude.md.example - PolarNet项目示例

## 验证配置

```bash
claude --version
claude config list
claude mcp list
```
