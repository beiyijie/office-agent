# Office Agent

<p align="center">
  <strong>A terminal-first office automation agent for email, documents, files, content analysis, and Tencent Meeting workflows.</strong>
</p>

<p align="center">
  <img alt="Python" src="https://img.shields.io/badge/Python-3.11+-3776AB?style=flat-square&logo=python&logoColor=white">
  <img alt="LangGraph" src="https://img.shields.io/badge/Workflow-LangGraph-111111?style=flat-square">
  <img alt="Tencent Meeting" src="https://img.shields.io/badge/Meeting-Tencent%20Meeting-0A84FF?style=flat-square">
  <img alt="License" src="https://img.shields.io/badge/License-MIT-green?style=flat-square">
</p>

<p align="center">
  <em>让邮件、文档、会议和通知从分散动作变成一条可执行工作流。</em>
</p>

## Table of Contents

- [Why This Project](#why-this-project)
- [Features](#features)
- [Workflow Examples](#workflow-examples)
- [Architecture](#architecture)
- [Project Structure](#project-structure)
- [Quick Start](#quick-start)
- [Configuration](#configuration)
- [Command Examples](#command-examples)
- [Tencent Meeting Integration Notes](#tencent-meeting-integration-notes)
- [Current Limitations](#current-limitations)
- [Roadmap](#roadmap)
- [Documentation](#documentation)
- [License](#license)

`Office Agent` 是一个面向真实办公场景的终端自动化 Agent。  
它不是简单的命令集合，而是基于状态流和节点编排，把“读文档、发邮件、建会议、写周报、做分析”这些动作串成可执行工作流。

## Why This Project

日常办公里最常见的问题不是“某个工具不会用”，而是流程被割裂：

- 邮件在邮箱里
- 文档在本地目录里
- 会议在会议平台里
- 总结、通知、汇报又要手工整理一次

`Office Agent` 的目标，就是把这些动作接成一条线。

## Features

### Email

- 查看最近邮件
- 搜索邮件
- 发送邮件
- 转发最近邮件
- 读取内容后继续交给其他节点处理

### Documents

- 读取 Word / Excel
- 生成 Word / Excel
- 生成周报、通知类文档
- 读取文档后继续发送或分析

### File Operations

- 列目录
- 读取文件
- 写入文件

### Content Analysis

- 分析 Word / Excel 内容
- 分析最近邮件内容
- 未配置 LLM 时自动降级为本地结果

### Tencent Meeting

- 创建腾讯会议
- 取消腾讯会议
- 创建后自动邮件通知
- 取消后自动邮件通知
- 支持 `dry_run` 模式，无商业版账号也能先跑通流程

## Workflow Examples

当前已经支持的典型工作流：

- `读取 weekly_report.docx 发给 test@example.com`
- `分析一下 weekly_report.docx 的内容`
- `总结最近的邮件然后生成周报文档`
- `创建明天下午3点到4点的腾讯会议并邮件通知 test@example.com`
- `取消明天下午3点的腾讯会议并邮件通知 test@example.com`

## Architecture

项目核心是一个轻量状态机：

```text
User Task
   |
   v
router_node
   |
   +--> email_node
   +--> document_node
   +--> file_ops_node
   +--> meeting_node
   +--> general_node
   |
   v
intermediate_results
   |
   v
next node / END
```

设计重点：

- `router_node` 负责识别意图并规划动作序列
- 各节点通过 `intermediate_results` 共享中间结果
- Agent 支持多步回流，而不是只执行单个命令

## Project Structure

```text
office_agent/
├── main.py
├── config.py
├── state.py
├── README.md
├── AGENT_INTRO.md
├── graph/
│   ├── agent.py
│   ├── edges.py
│   └── nodes.py
└── tools/
    ├── document.py
    ├── email.py
    ├── file_ops.py
    └── tencent_meeting.py
```

## Quick Start

### 1. Clone and install

```powershell
cd office_agent
pip install -r requirements.txt
```

### 2. Create config

```powershell
Copy-Item config.example.json config.json
```

### 3. Start the agent

```powershell
python main.py
```

## Configuration

支持两种配置方式，优先级如下：

1. 环境变量
2. `config.json`

安全说明：

- `config.json` 已在 `.gitignore` 中屏蔽，不会被提交到 GitHub
- 仓库内应只保留 `config.example.json` 作为公开模板
- 上传前不要把真实密钥写进 README、示例截图或 issue

### QQ Mail

```json
{
  "qq_email": {
    "email": "your_qq@qq.com",
    "auth_code": "your_qq_auth_code"
  }
}
```

### Minimax

```json
{
  "minimax": {
    "api_key": "your_minimax_api_key"
  }
}
```

### Tencent Meeting

```json
{
  "tencent_meeting": {
    "access_token": "your_tencent_meeting_access_token",
    "user_id": "your_operator_user_id",
    "dry_run": true
  }
}
```

说明：

- `dry_run=true` 时，会生成模拟会议结果，但邮件通知和整条工作流仍然会照常执行
- 这适合没有腾讯会议商业版账号时先开发和演示

### Environment Variables

```powershell
$env:QQ_EMAIL="your_qq@qq.com"
$env:QQ_AUTH_CODE="your_qq_auth_code"
$env:MINIMAX_API_KEY="your_minimax_api_key"
$env:TENCENT_MEETING_ACCESS_TOKEN="your_tencent_meeting_access_token"
$env:TENCENT_MEETING_USER_ID="your_operator_user_id"
$env:TENCENT_MEETING_DRY_RUN="true"
```

## Command Examples

```text
查看最近5封邮件
发送邮件给 manager@company.com 主题周报 内容本周任务已完成
读取 weekly_report.docx 发给 test@example.com
分析一下 weekly_report.docx 的内容
总结最近的邮件然后生成周报文档
创建明天下午3点到4点的腾讯会议并邮件通知 test@example.com
取消明天下午3点的腾讯会议并邮件通知 test@example.com
```

## Tencent Meeting Integration Notes

当前腾讯会议接入采用“先跑通工程闭环，再接真实 API”的思路：

- 已实现 `create_meeting`
- 已实现 `cancel_meeting`
- 已实现会议结果转邮件通知
- 已支持 `dry_run`

这意味着：

- 没有商业版账号也可以先完成本地验证
- 有真实开放平台能力后，可切换到正式 API 调用

## Current Limitations

- 腾讯会议真实 API 尚未实际联调
- 自然语言时间解析目前支持常见表达，但不是完整时间语义引擎
- 会议取消在没有真实 `meeting_id` 时，会按解析结果生成模拟对象再通知
- 邮件转发目前默认取最近一封邮件

## Roadmap

- 支持修改会议时间并自动邮件通知
- 支持基于 `meeting_id / meeting_code` 精确取消会议
- 支持保存最近一次会议上下文，处理“取消刚才那个会议”
- 接入飞书、企业微信、Outlook 等更多办公平台
- 增加自动化测试和回归用例

## Documentation

- 项目介绍见 [AGENT_INTRO.md](./AGENT_INTRO.md)
- 配置模板见 [config.example.json](./config.example.json)

## License

本项目当前使用 [MIT License](./LICENSE)。
