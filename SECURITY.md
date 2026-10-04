# Security and private data / 安全与私有数据

This project reads mailbox metadata and saves sensitive local artifacts. It is not a credential vault.

- Never attach real mailbox reports, subjects, addresses, folder names, access/refresh tokens, device codes or `.env` files to public issues.
- Keep token caches and reports in a private directory with appropriate OS permissions, outside cloud-sync/public/shared folders. Token cache files are plaintext. Windows requires suitable ACLs; `chmod` is not a Windows ACL configuration mechanism.
- Removing a cache does not revoke tokens. Use your Microsoft account/organization's application access controls if credentials have been exposed.
- CSV formula-like strings are escaped for spreadsheet viewing; re-saving/converting a CSV can change that protection. JSON reports preserve original values.
- The application exposes GET-only Graph reads, validates pagination hosts and uses no mail/rule write APIs. Authentication uses Microsoft's standard login/token endpoints. Classic Outlook may perform its own background sync.
- Use **GitHub private vulnerability reporting** if the repository has it enabled. Otherwise open a public issue requesting a private contact channel without exploit details or private data.

本工具会读取邮箱元数据并生成敏感本地文件，不提供凭据保险库功能。

- 不要在公开 issue 中附上真实邮箱报告、主题、地址、文件夹名称、访问/刷新令牌、设备登录码或 `.env`。
- 缓存与报告应存放在私有且权限恰当的目录，避免云同步、公共或共享目录。缓存为明文；Windows 需设置合适的 ACL，`chmod` 不能替代 ACL。
- 删除缓存不等于吊销令牌。如凭据泄露，应使用微软账号/组织提供的应用访问管理功能处理。
- CSV 会转义类似公式的文本，但另存或转换可能改变防护效果；JSON 保留原值。
- 工具仅暴露 Graph GET 读取，对分页域名做校验，不调用邮件或规则写接口。认证通过微软标准登录/令牌接口；经典 Outlook 自身可能进行后台同步。
- 若仓库启用 GitHub 私密漏洞报告，请通过该渠道提交。否则先发不含漏洞细节或私有数据的 issue 请求私密联系方式。
