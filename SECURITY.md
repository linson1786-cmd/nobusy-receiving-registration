# Security Policy

## 数据边界

本仓库只存放 NoBusy 别虾忙｜AI 管理实战的 WorkBuddy Skill 源码、说明文档和示例配置。

禁止提交：

- 密码、Token、Cookie、API Key、邮箱授权码；
- `config.py`、`.env`、证书、密钥文件；
- 真实收货登记表、Excel 台账、PDF/OFD 附件；
- 真实个人、真实组织、真实审批系统或真实业务数据；
- WorkBuddy 内部目录、部署备份和本地缓存。

## 配置规则

- 本地配置保留在本机；
- 仓库只保留示例、模板和通用规则；
- 示例数据必须使用虚构内容；
- 发布前必须执行敏感信息扫描。

## 发布前检查

```bash
git status --short
python3 -m py_compile scripts/receiving-registration/*.py
rg -n "真实姓名|真实组织|真实路径|password|token|secret|auth" .
find . -name ".DS_Store" -o -name "__pycache__" -o -name "config.py" -o -name "*.zip"
```

命中 `config.py`、压缩包、真实数据或敏感信息时，不允许提交或发布。

