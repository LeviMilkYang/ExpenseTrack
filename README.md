# Telegram Expense Bot

这个项目会轮询 Telegram Bot API，把白名单用户发来的文本消息交给 Codex 判断是否属于记账内容；如果是，就写入 `expense.xlsx`，否则回复忽略提示。

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

所有本地私密配置和 bot 运行态都保存在 `bot_runtime/telegram_bot_config.json`，不要提交到 Git。

当前这个文件里至少会用到这些字段：

- `token`：Telegram bot token
- `offset`：Telegram `getUpdates` 的消费游标
- `report_period`：当前报表月份缓存
- `pending_restart_notice`：重启确认消息的待发送状态
- `allowed_username`：白名单 Telegram 用户名
- `project_dir`：项目根目录绝对路径，用来推导 `expense.xlsx` 和 `expense_report.xlsx`
- `allowed_categories`：合法收支分类集合
- `payment_channels`：合法支付渠道集合，daemon 启动时一次性加载到内存
- `default_payment_channel`：默认支付渠道，必须是 `payment_channels` 里的一个值

## 核心特性

- **自动识别记账内容**：白名单用户发送自然语言消息后，机器人会判断是否属于记账内容，并写入 `expense.xlsx`。
- **失败时保留原始记录**：当 AI 无法正确解析时，机器人会把原文以待确认记录写入账本，避免漏记。
- **支持作废与报表联动**：支持作废上一条或指定消息对应的记录；被作废的数据不会进入统计报表。

## 当前行为

- **智能判断**：白名单用户发来的文本都会交给 Codex 处理。
- **快捷指令**：
    - `重启`：直接触发重启脚本，完成后尝试发送确认。
    - `作废`：支持直接回复某条记录进行精确作废，或作废上一条记录。
    - `预算`：读取 `expense.xlsx` 的 `预算` sheet，显示当月预算项的已用额度和剩余额度；`Fixed` 标记为真时，该项不在预算明细中展示。
- **自动补全**：若消息无明确日期时间，优先使用 Telegram 消息的时间戳。
- **默认支付渠道**：若消息没有明确提到支付渠道，使用配置文件中的 `default_payment_channel`。
- **时区声明**：消息里可以显式声明当前所在时区，例如 `UTC+0`、`UTC-5`、`UTC+05:30`。默认按 `UTC+08:00` 处理。
- **报表联动**：`Status = 作废` 的记录不会计入统计，报表每月自动刷新。
- **认证失败快速退出**：若 Telegram API 返回 `401 Unauthorized`，守护进程会记录致命认证错误并直接退出，避免持续刷日志。

## 运行与维护

以下脚本必须在 WSL/Linux 环境里执行；Windows 侧不能可靠判断 WSL 里的 daemon 进程状态。

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

- **WSL/Linux**: `python3`, `codex` CLI (已登录)
- **Windows**: `python`, `openpyxl`
