# 收货信息登记 Skill

> **版本**: V1.0  
> **创建日期**: 2026-06-23  
> **项目归属**: NoBusy 别虾忙｜AI 管理实战  
> **负责人**: Linson  
> **License**: MIT

---

## 项目简介

本项目为 WorkBuddy 的 Skill 插件，用于仓储部门收货信息的交互式采集、结构化记录和自动化登记。通过对话采集收货信息，自动生成编号，写入本地 Excel 文件，替代传统手工填写表单/Excel 的方式。支持单物品/多物品批量登记、首次表初始化、已有记录补充修正。

## 功能列表

| 功能 | 脚本 | 说明 |
|------|------|------|
| 收货登记 | `receiving_excel.py add` | 交互采集收货信息，自动编号写入 Excel |
| 表初始化 | `receiving_excel.py init` | 首次使用创建标准化收货登记表（18列表头） |
| 编号生成 | `receiving_excel.py next-number` | 查询当日已有记录数，生成 RG-YYYYMMDD-NNN 编号 |
| 记录查询 | `receiving_excel.py query` | 按编号精确/日期供应商模糊查询 |
| 记录更新 | `receiving_excel.py update` | 按编号定位更新，保留单元格样式 |
| 主体列表 | `receiving_excel.py list-companies` | 从注册表获取业务主体列表 |
| 部署工具 | `deploy.py` | 版本对比、备份、自动部署到 ~/.workbuddy/skills/ |

---

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 部署 Skill

```bash
cd scripts
python3 deploy.py
```

部署工具会自动将 Skill 文件同步到 `~/.workbuddy/skills/receiving-registration/`。

### 3. 初始化收货表

在 WorkBuddy 对话中说"初始化收货表"，按引导提供主体名称和保存路径。

### 4. 开始登记

在 WorkBuddy 对话中说"收货登记"或"到货了登记一下"即可触发登记流程。

---

## 项目结构

```
receiving-registration/
├── .gitignore              # Git 忽略规则
├── LICENSE                 # MIT 许可证
├── requirements.txt        # Python 依赖
├── README.md              # 本文件
├── SKILL.md               # Skill 触发词文档
├── V1.0-交付说明.md       # 版本交付说明
├── 项目管理体系说明.md     # 项目管理体系
├── references/            # 参考文档（工作目录）
│   ├── field_schema.md    # 字段定义、校验规则
│   └── company_registry.md # 主体注册表
├── scripts/               # 核心脚本（开发工作目录）
│   ├── VERSION            # 当前版本 V1.0
│   ├── CHANGELOG.md       # 变更日志
│   ├── deploy.py          # 部署工具
│   └── receiving_excel.py # Excel 核心操作脚本
├── 源文件/                # 版本源码归档
│   └── V1.0/
├── 交付包/                # 对外发布包
│   └── V1.0/
├── 版本归档/              # 历史版本压缩包
├── 项目管理/              # 项目管理文档
│   ├── 版本发布记录/
│   ├── 需求与计划/
│   ├── 测试记录/
│   └── 使用反馈/
└── Templates/             # 模板文件
```

---

## 版本管理

本项目采用 V1.0 / V1.1 版本号格式。使用 deploy.py 自动检测更新并备份旧版本。

### 发布新版本

1. 在 `scripts/` 中修改代码
2. 更新 `scripts/VERSION` 文件
3. 更新 `scripts/CHANGELOG.md`
4. 同步到 `源文件/V{version}/` 和 `交付包/V{version}/`
5. 创建归档到 `版本归档/`
6. 运行 `python3 scripts/deploy.py` 部署

---

## 注意事项

- 用户数据（`~/收货登记/*.xlsx`）独立于代码，更新不受影响
- 各业务主体使用独立的 Excel 文件，互不混用
- 脚本依赖 openpyxl 库，首次使用前需安装

---

*本项目为 NoBusy 别虾忙｜AI 管理实战 Skill 工具。*
