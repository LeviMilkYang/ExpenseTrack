---
name: expense-bot-maintainer
description: 维护 Telegram 记账机器人时使用。适用于排查消息处理、Excel 写入、配置加载、日志问题，以及新增字段或记账能力时快速建立上下文。
---

# Expense Bot Maintainer

## 目标

这个仓库的核心不是 Web 服务，而是一个常驻的 Telegram 轮询 daemon。后续模型进入仓库后，优先按“消息从哪里进来，在哪里被 LLM 解析，最后如何写进 Excel”这条链路理解代码，不要先从 README 展开。

## 先看哪些文件

- `bot_runtime/telegram_expense_daemon.py`
  入口和主流程。负责：
  - 启动时读取 `bot_runtime/telegram_bot_config.json`
  - 轮询 Telegram `getUpdates`
  - 调 bridge 生成 prompt
  - 调 `codex exec` 拿结构化 JSON
  - 调 bridge/apply 和 Excel 写入
  - 发回复、记日志、处理重启和作废

- `bot_runtime/telegram_codex_bridge.py`
  LLM 前后处理层。负责：
  - 生成 prompt
  - 把 daemon 注入的运行时配置透传给 prompt
  - 把模型输出从 `tool_call` 标准化成真实可执行的 Excel tool payload
  - 给 `append_record` 补默认值
  - 决定写入哪个 sheet

- `bot_runtime/append_excel_entry.py`
  Excel 写入和记录规范化逻辑。这里定义：
  - `EXPECTED_HEADERS`
  - `normalize_record`
  - 表头迁移逻辑
  - 排序逻辑
  - 通过 helper 调 Windows Python + openpyxl 实际写表

- `bot_runtime/excel_tools.py`
  Excel tool 调度入口。负责：
  - 用统一 JSON payload 调用 Excel 能力
  - 暴露工具清单 `describe`
  - 把 `append_record` / `invalidate_record` / `read_record` / `invalidate_last_record` 统一成后续 agent 可直接消费的接口

- `bot_runtime/tmp_excel_helper/excel_append_helper.py`
  运行时 helper 的落地副本。它经常和 `append_excel_entry.py` 里的 `HELPER_CODE_OPENPYXL` 逻辑保持一致。排查 Excel 行为时，这两个文件都要一起看。

- `bot_runtime/telegram_record_schema.py`
  `codex exec --output-schema` 用的输出 schema。当前主链路要求模型返回 `tool_call` 而不是直接返回 `record`。新增账本字段或新增 tool 时，这里通常也要同步改。

## 真实数据流

1. Telegram 消息进入 `telegram_expense_daemon.py`
2. daemon 组装 envelope
3. daemon 用内存中的配置构造 `runtime_config`
4. `telegram_codex_bridge.py prompt` 生成 prompt
5. daemon 调 `codex exec`
6. `telegram_codex_bridge.py apply` 把模型输出转成标准化 `tool_call`
7. `excel_tools.py` / `append_excel_entry.py` 调 helper 写入 Excel
8. helper 追加、排序、保存，再返回最终行号
9. daemon 回 Telegram 发送确认消息，并把索引写到月度 index 文件

补充：如果后续不是 daemon 直接调 Python 函数，而是 LLM/agent 走结构化 tool 调用，则应优先从 `excel_tools.py` 进入，再由它调用 `append_excel_entry.py` 中的底层能力。

## 配置规则

- 配置文件是 `bot_runtime/telegram_bot_config.json`
- 主链路要求：启动时一次性加载到内存，不做热加载
- 不要在每次 LLM 调用时重新读取配置文件，除非用户明确要求改回实时读取
- 新增“受配置驱动的枚举项”时，优先放到这个 JSON，而不是写死在 prompt 或 Python 常量里
- 支付渠道默认值也走配置：`default_payment_channel` 必须放在 `bot_runtime/telegram_bot_config.json`，并且要属于 `payment_channels`

## 记录结构规则

- 账本字段以 `append_excel_entry.py` 的 `EXPECTED_HEADERS` 为准
- LLM 输出结构以 `telegram_record_schema.py` 为准
- bridge 的 prompt 说明、tool_call 结构和默认值逻辑也要同步
- 新增字段时，通常至少要改这 4 处：
  - `telegram_record_schema.py`
  - `telegram_codex_bridge.py`
  - `append_excel_entry.py`
  - helper 代码/文件

- 如果新增的是一个新的 Excel tool，而不是单纯新增 record 字段，则通常还要同步：
  - `excel_tools.py`
  - `telegram_codex_bridge.py` prompt 和 apply 逻辑
  - `telegram_record_schema.py`

## Excel 相关高风险点

- 这个项目不是直接用 WSL Python 写最终 Excel；默认路径是：
  WSL Python -> 生成 helper -> `powershell.exe` -> Windows Python/openpyxl
- `append_excel_entry.py` 中的 `HELPER_CODE_OPENPYXL` 是实际执行代码来源
- `bot_runtime/tmp_excel_helper/excel_append_helper.py` 是辅助副本，提交时应保持一致，避免后续排查时上下文错位
- 写入后会按日期时间重新排序，任何新增字段都必须保证“整行搬运”时不会串值
- 之前出现过的 bug：
  - 排序回写时用 `worksheet.cell(..., value=val)`，当 `val is None` 时不会真正清空旧值，导致空字段残留并污染其他记录
  - 修复方式是先取 `cell`，再显式 `cell.value = val`

## 排查顺序

### 用户说“没反应”

- 先看 `bash bot_runtime/status_bot.sh`
- 再看 `bot_runtime/logs/$(date +%Y-%m).log`
- 判断是：
  - daemon 没跑
  - Telegram API 网络抖动
  - LLM 失败走了 fallback
  - 回复发送失败但写账成功

### 用户说“识别错了”

- 先看日志里的 `message_received`、`message_applied`
- 再看 bridge prompt 是否拿到了正确配置
- 再看 schema 是否限制过严或默认值覆盖了模型输出

### 用户说“Excel 被写坏了”

- 优先怀疑 `append_excel_entry.py` 和 helper
- 特别检查：
  - 表头迁移
  - 排序回写
  - 空值写回
  - 返回行号是否和排序后的 ID 匹配

## 开发约束

- 除非用户明确要求，否则不要引入热加载
- 除非用户明确要求，否则不要把配置枚举重新写死回代码
- 修改 Excel 字段时，必须考虑旧表头兼容
- 修改 daemon 行为时，优先保持日志字段稳定，避免排查链路断掉
- 不要顺手重构大段无关代码；这个仓库更看重稳定性而不是“风格统一”

## 文档更新规则

- 新功能开发后：
  - `skill.md` 的更新保持在技术架构维度，只记录后续模型排查和开发真正需要的上下文、约束、坑点、链路
  - `README.md` 的更新保持在功能维度，只描述用户可见行为、运行方式、配置项
- 除非是明显的新特性、用户可见行为变化或运维方式变化，否则不要随意给 `README.md` 增加内容
- 如果只是修复内部实现、重构、兼容旧表头、修 helper、改日志结构，通常更应该更新 `skill.md` 而不是 README

## 提交前检查

- `python3 -m py_compile` 至少覆盖被改的 Python 文件
- 如果改了 Excel 写入逻辑，优先构造最小临时 workbook 做复现测试
- 如果改了配置驱动逻辑，确认 daemon 主链路走的是内存配置，不是每次现读文件
- 提交到 GitHub 前，检查不要把本地私密配置或运行态脏数据带进去
