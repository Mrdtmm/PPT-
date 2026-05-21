# MCP Servers 配置模板

本文档列出 Claude Code 可用的官方 MCP (Model Context Protocol) 服务器及其配置方法。

## 安装 MCP Server

在 Claude Code 中使用 MCP server 需要通过官方插件市场安装：

```bash
# 查看可用插件
npx skills list

# 安装 MCP server
npx skills add @anthropic/claude-plugin-{name}
```

## 官方 MCP Servers

| Server | 用途 | 安装命令 |
|--------|------|----------|
| github | GitHub API 操作（issues, PRs, repos） | npx skills add @anthropic/claude-plugin-github |
| gitlab | GitLab API 操作 | npx skills add @anthropic/claude-plugin-gitlab |
| asana | 项目管理 | npx skills add @anthropic/claude-plugin-asana |
| linear | 问题跟踪 | npx skills add @anthropic/claude-plugin-linear |
| discord | Discord 机器人 | npx skills add @anthropic/claude-plugin-discord |
| telegram | Telegram 机器人 | npx skills add @anthropic/claude-plugin-telegram |
| firebase | Firebase 集成 | npx skills add @anthropic/claude-plugin-firebase |
| context7 | 代码库上下文 | npx skills add @anthropic/claude-plugin-context7 |
| greptile | 代码搜索 | npx skills add @anthropic/claude-plugin-greptile |
| playwright | 浏览器自动化 | npx skills add @anthropic/claude-plugin-playwright |
| terraform | Terraform IaC | npx skills add @anthropic/claude-plugin-terraform |
| serena | 通用助手 | npx skills add @anthropic/claude-plugin-serena |

## 自定义 MCP Server

如果需要添加自定义 MCP server，可以在 settings.json 中配置：

```json
{
  "mcpServers": {
    "my-server": {
      "command": "npx",
      "args": ["-y", "@my/mcp-server"],
      "env": {
        "API_KEY": "your-api-key"
      }
    }
  }
}
```

## 常用 MCP 配置示例

### GitHub MCP
```json
{
  "mcpServers": {
    "github": {
      "command": "npx",
      "args": ["-y", "@anthropic/claude-plugin-github"],
      "env": {
        "GITHUB_TOKEN": "ghp_xxxxxxxxxxxx"
      }
    }
  }
}
```

### 本地PPTX MCP (pptx-mcp)
```json
{
  "mcpServers": {
    "pptx-mcp": {
      "command": "python",
      "args": ["E:/CODE/zhukefu/pptx-mcp/server.py"],
      "env": {}
    }
  }
}
```

## 环境变量配置

创建 .env.template 文件：

```bash
# GitHub
GITHUB_TOKEN=ghp_xxxxxxxxxxxx

# 其他 API Keys
OPENAI_API_KEY=sk-xxxxxx
ANTHROPIC_API_KEY=sk-ant-xxxxxx
```

## 验证 MCP 配置

```bash
# 检查 MCP server 状态
claude mcp list

# 测试连接
claude mcp test {server-name}
```
