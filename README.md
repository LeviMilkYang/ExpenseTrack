# Telegram Expense Bot

这个项目会轮询 Telegram Bot API，把白名单用户发来的文本消息交给 Gemini 判断是否属于记账内容；如果是，就写入 `expense.xlsx`，否则回复忽略提示。

## 项目结构

- 运行目录：`bot_runtime/`
- 主守护程序：`bot_runtime/telegram_expense_daemon.py`
- 启动脚本：`bot_runtime/start_bot.sh`
- 重启脚本：`bot_runtime/restart_bot.sh`
- 停止脚本：`bot_runtime/stop_bot.sh`
- 状态脚本：`bot_runtime/status_bot.sh`
- 生成报表：`bot_runtime/update_report.sh`
- 历史日志：`bot_runtime/logs/` (按月分割)
- 历史索引：`bot_runtime/indexes/` (按月分割)
- Excel 主文件：`expense.xlsx`
- 统计报表：`expense_report.xlsx`

## 本地私有配置

所有本地私密配置都保存在 `bot_runtime/telegram_bot_state.json`，不要提交到 Git。

当前这个文件里至少会用到这些字段：

- `token`：Telegram bot token
- `offset`：Telegram `getUpdates` 的消费游标
- `allowed_username`：白名单 Telegram 用户名
- `project_dir`：项目根目录绝对路径，用来推导 `expense.xlsx` 和 `expense_report.xlsx`

## 核心特性

### 1. 交互与反馈
- **即时响应**：收到记账消息后立即回复“正在处理...”，提升交互确定性。
- **详细结果**：处理完成后告知记账分类、金额及在 Excel 中的具体行号。

### 2. 高可用性与容错
- **AI 智能重试**：当 Gemini 处理失败（网络波动或解析错误）时，会自动进行 3 次重试（采用指数退避算法：1s, 2s, 4s）。
- **无损兜底录入**：如果 AI 在重试后依然彻底失败，机器人会自动将原始消息记入 Excel：
    - 金额设为 `0`，分类设为 `未分类`。
    - 备注标记 `[AI失败兜底]` 并包含完整原文。
    - 标记为 `NeedConfirm` (待确认)，并提示用户手动核对行号。

### 3. 数据与文件管理
- **按月分割存储**：
    - **日志**：存放在 `bot_runtime/logs/YYYY-MM.log`，方便定期清理。
    - **索引**：存放在 `bot_runtime/indexes/YYYY-MM.json`，确保精确作废功能长期可靠。
- **后台日志去重**：守护进程后台运行时只写月日志文件，不再同时通过 `stdout` 重复落盘。
- **环境隔离**：调用 `gemini` 时自动清理 IDE 相关环境变量，确保在 CLI 环境下稳定运行。

## 当前行为

- **智能判断**：白名单用户发来的文本都会交给 Gemini 处理。
- **快捷指令**：
    - `重启`：直接触发重启脚本，完成后发送确认。
    - `作废`：支持直接回复某条记录进行精确作废，或作废上一条记录。
- **自动补全**：若消息无明确日期时间，优先使用 Telegram 消息的时间戳。
- **报表联动**：`NeedConfirm = 作废` 的记录不会计入统计，报表每月自动刷新。
- **认证失败快速退出**：若 Telegram API 返回 `401 Unauthorized`，守护进程会记录致命认证错误并直接退出，避免持续刷日志。

## 运行与维护

启动/停止/状态/重启：
```bash
bash bot_runtime/start_bot.sh
bash bot_runtime/stop_bot.sh
bash bot_runtime/status_bot.sh
bash bot_runtime/restart_bot.sh
```

查看当前月日志：
```bash
tail -f bot_runtime/logs/$(date +"%Y-%m").log
```

## 依赖

- **WSL/Linux**: `python3`, `gemini` CLI (已登录)
- **Windows**: `python`, `openpyxl`
