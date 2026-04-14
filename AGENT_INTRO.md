# Office Agent Overview

## What It Is

`Office Agent` 是一个以终端为入口的办公自动化 Agent，用来把常见办公动作组织成连续工作流，而不是一个个孤立命令。

它目前聚焦五类能力：

- 邮件处理
- 文档读写
- 文件操作
- 内容分析
- 腾讯会议与邮件通知联动

## Core Idea

项目的重点不是“工具多”，而是“工具之间能接起来”。

例如：

- 读取文档后直接发邮件
- 总结最近邮件后生成周报
- 创建腾讯会议后自动给参会人发通知
- 取消腾讯会议后自动发取消邮件

这类流程都通过统一状态结构和节点回流机制实现。

## Technical Shape

核心结构由三部分组成：

### 1. State

`AgentState` 保存：

- 当前任务
- 对话消息
- 工具结果
- 下一步动作
- 已完成动作
- 中间结果 `intermediate_results`

### 2. Router

`router_node` 负责做动作规划，例如：

- `document_node -> email_node`
- `email_node -> document_node`
- `meeting_node -> email_node`
- `document_node -> general_node`

### 3. Nodes

每个节点负责一类业务：

- `email_node`
- `document_node`
- `file_ops_node`
- `meeting_node`
- `general_node`

节点之间通过 `intermediate_results` 共享内容、路径、摘要、会议信息等数据。

## Why Tencent Meeting First

在当前阶段，腾讯会议比完整日历系统更适合作为第一步扩展：

- 场景更明确
- 工程边界更清晰
- 很适合和邮件通知联动

同时项目也支持 `dry_run`：

- 没有商业版账号也能先本地验证完整工作流
- 后续只需补真实凭证即可切到正式 API

## Suitable Use Cases

- 个人办公自动化
- 小团队内部效率工具
- 需要邮件、文档、会议联动的轻量场景
- 国内办公环境下的原型系统

## Current Status

目前项目已经具备一个可演示、可继续扩展的基础版本：

- 支持多步工作流
- 支持腾讯会议创建/取消后发邮件
- 支持文档读取、生成与分析
- 支持邮件内容读取、总结与发送

它已经不是一个空壳原型，而是一个可以继续往真实办公助手方向推进的骨架。

## Suggested Next Step

如果后续继续演进，优先级建议是：

1. 修改会议时间并通知
2. 更强的会议上下文记忆
3. 更精确的自然语言时间解析
4. 飞书 / 企业微信 / Outlook 接入
5. 自动化测试体系
