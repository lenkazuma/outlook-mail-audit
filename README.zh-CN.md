# Outlook Mail Audit · Outlook 邮箱只读盘点

[English](README.md) · [安全说明](SECURITY.md)

用于个人整理 Outlook 邮箱的小工具：盘点文件夹、现有收件箱规则和未分类发件人，生成本地 CSV/JSON 报告，并给出供手工审阅的规则建议。工具不会创建或执行规则，不会移动、发送或删除邮件。

这个仓库用于归档个人实际用过的工具，不是微软官方产品，也不是完整的邮箱合规审计系统。

## 两种读取方式

| 方式 | 前提 | 登录 | 扫描范围 |
| --- | --- | --- | --- |
| Microsoft Graph | Python 3.11+、微软邮箱、自行注册的应用 | 设备代码交互登录 | 当前登录账号 |
| 本机经典 Outlook | Windows、Python 3.11+、已配置的经典桌面 Outlook | 使用已有 Outlook 配置 | 配置中挂载的全部邮箱存储 |

新版 Outlook for Windows 不提供经典 COM 接口；没有经典版时可使用 Graph。查看帮助和运行离线测试只需 Python，实际扫描时才加载相应依赖。

## 安装

```powershell
git clone https://github.com/lenkazuma/outlook-mail-audit.git
cd outlook-mail-audit
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

macOS/Linux 使用 `source .venv/bin/activate` 激活环境，且仅支持 Graph 方式。

### 方式 A：Windows 经典 Outlook

先打开经典 Outlook，确认当前使用的是要扫描的配置，然后运行：

```powershell
python -m pip install -r requirements-outlook.txt
python scan_outlook_local_readonly.py --days 365 --max-messages-per-folder 5000
```

会读取该配置挂载的所有邮箱存储，可能包括共享邮箱和存档。结果写入 `output_local/`。这条路径不需要注册 Azure/Entra 应用。

### 方式 B：Microsoft Graph

1. 在 [Microsoft Entra 应用注册](https://entra.microsoft.com/) 创建自己的应用；支持的账号类型按实际需要选择，需要个人微软账号时要包含该类型。
2. 在 **Authentication / 身份验证** 中启用 **Allow public client flows / 允许公共客户端流**。本工具使用设备代码登录，不使用客户端密钥。
3. 添加 Microsoft Graph **委托权限**：`User.Read`、`Mail.Read`、`MailboxSettings.Read`。是否需要管理员同意取决于组织策略。
4. 安装依赖并复制配置模板：

```powershell
python -m pip install -r requirements.txt
Copy-Item .env.example .env
```

在本地 `.env` 中填写：

```dotenv
GRAPH_CLIENT_ID=your_app_client_id
GRAPH_TENANT=common
GRAPH_SCOPES=User.Read Mail.Read MailboxSettings.Read
```

把占位符改为自己的应用客户端 ID。`common` 配合应用所支持的账号类型使用；组织有要求时使用自己的租户 ID。MSAL 自动添加保留的 OpenID scopes，不需在配置中填写 `offline_access`、`openid` 或 `profile`。

```powershell
python scan_readonly.py --days 365 --max-messages-per-folder 5000
```

按照设备代码提示在微软页面登录，核对页面上的账号和应用名称。结果写入 `output/`。不同账号应使用不同缓存，例如 `--token-cache .token_cache_work.bin`；缓存含多个账号时工具不会自行选择。

微软官方参考：[MSAL 设备登录](https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens)、[读取收件箱规则及权限](https://learn.microsoft.com/en-us/graph/api/mailfolder-list-messagerules?view=graph-rest-1.0)、[邮件文件夹标识](https://learn.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0)。

## 参数和结果含义

```powershell
python scan_readonly.py --help
python scan_outlook_local_readonly.py --help
```

两个入口均支持 `--days`（默认 365）、`--max-messages-per-folder`（默认 5000）、`--suggest-threshold`（默认 5）和 `--output-dir`。数值必须为正整数。Graph 额外支持 `--token-cache`。

- “未整理 / unsorted”仅表示直接位于默认收件箱且没有分类标签，不代表规则失效。
- 根据文件夹 ID 识别收件箱，支持中文等本地化名称；收件箱下的子文件夹不算在内。
- 日期范围根据接收时间计算。统计受时间和每文件夹条数上限影响，并非全部历史邮件；文件夹总数与扫描样本数可能不同。
- 规则建议只写入 JSON，需自行审阅。Graph 会检查已有启用规则中的发件人关键字，但不会完整模拟条件、例外、先后顺序和实际执行。本机 Outlook 规则摘要较简略，也不会排除已被现有规则覆盖的发件人。

## 本地报告

| 文件 | 内容 |
| --- | --- |
| `mailbox_profile.json` | 扫描时间、账号/配置和计数；本机方式另含部分跳过存储/文件夹的错误 |
| `folders.csv` | 文件夹名称与计数 |
| `rules_inventory.csv` | 现有规则摘要 |
| `all_messages_metadata.csv` | 样本邮件主题、发件人、日期、分类和标识 |
| `unsorted_inbox_messages.csv` | 直接位于收件箱且无分类的邮件 |
| `unsorted_senders.csv` | 上述邮件的发件人计数 |
| `proposed_rules.json` | 供手工审阅的规则建议 |
| `folders.json`、`rules.json` | 仅 Graph：原始文件夹/规则记录 |

不导出邮件正文或附件内容；规则定义本身可能包含敏感的条件文字。空 CSV 是空文件。CSV 文本以公式字符开头时会添加英文单引号，降低电子表格将其当作公式执行的风险；JSON 原值不变。由于 API 不同，两种方式的 CSV 字段不完全一致。

[examples/synthetic](examples/synthetic) 是未连接任何邮箱而制作的虚构示例。真实邮箱数据不应提交到仓库。

## 隐私与只读边界

邮箱 API 仅使用 GET。Graph 权限限定为前述三项只读权限；请求和分页限定在 `https://graph.microsoft.com/v1.0/`，禁止重定向。登录流程会联系微软，并写入本地令牌缓存。本机方式读取已有 Outlook 配置并写报告；Outlook 自身仍可能进行正常后台同步。

Git 已忽略 `.env`、令牌缓存、默认输出目录和报告文件名。**自定义输出位置仍需在提交前自行核对。** 报告和缓存应存放在私有本地磁盘，避免公共、共享或自动同步的目录。缓存是明文，不是加密文件；脚本设置的 POSIX 文件权限不能替代 Windows 文件夹 ACL。删除缓存仅清除本工具本地登录状态，不会吊销已经签发的令牌。

## 离线测试

```powershell
python -m unittest discover -s tests -v
```

测试只使用虚构数据和模拟对象，不打开 Outlook、不登录微软、不联网或修改邮件。CI 在 Windows/Linux、Python 3.11/3.13 上运行同样的测试。

## 当前状态与待完善事项

已补齐公开归档基础：中英文说明、MIT 许可证、分离依赖、虚构示例、CSV 公式防护、本地化收件箱识别、Graph 域名限制、离线测试。

仍待完成：

- [ ] 在代表性环境验证全新 Graph 登录与经典 Outlook 实际扫描（离线 CI 未验证）。
- [ ] Graph 限流退避重试、长时间扫描中的令牌刷新。
- [ ] 本机方式指定邮箱存储，以及更完整的 Exchange 地址到 SMTP 地址转换。
- [ ] 更完整地分析规则覆盖与例外，并逐项记录读取不完整的本机规则。
- [ ] 结构化进度/错误报告和依赖更新检查。

当前 Graph 遇到未处理的错误（包括限流或令牌过期）会停止；失败后不要把旧报告当成刚完成的完整扫描。本机 Outlook 可能返回 Exchange 内部发件人地址，部分受保护文件夹或规则属性可能无法读取。规则建议始终需要人工核对。

## 许可证

[MIT](LICENSE)，版权所有 2026 Zhenzhen Liu。
