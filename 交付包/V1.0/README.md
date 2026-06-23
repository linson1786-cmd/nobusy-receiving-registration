# 收货信息登记 Skill

> **版本**: V1.0  
> **交付日期**: 2026-06-23  
> **项目归属**: NoBusy 别虾忙｜AI 管理实战  
> **License**: MIT

## 简介

仓储部门收货信息登记自动化工具。通过对话采集收货信息并写入本地 Excel，支持单/多物品批量登记、表初始化、记录补正。

## 快速开始

1. 安装依赖：`pip install openpyxl`
2. 复制本目录到 `~/.workbuddy/skills/receiving-registration/`
3. 在 WorkBuddy 对话中说"初始化收货表"开始使用

## 命令速查

| 触发词 | 功能 |
|--------|------|
| "收货登记""到货了登记一下" | 新增收货登记 |
| "初始化收货表""建一个收货登记表" | 创建标准化 Excel 表 |
| "修改收货记录""补充收货信息" | 查询并更新已有记录 |

## 文件清单

| 文件 | 说明 |
|------|------|
| `SKILL.md` | Skill 触发词文档 |
| `V1.0-交付说明.md` | 详细使用说明 |
| `scripts/receiving_excel.py` | Excel 核心操作（6个子命令） |
| `scripts/deploy.py` | 部署工具 |
| `references/field_schema.md` | 字段定义与校验规则 |
| `references/company_registry.md` | 主体注册表 |

---

*NoBusy 别虾忙｜AI 管理实战*
