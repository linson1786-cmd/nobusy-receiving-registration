#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
收货信息登记 Skill 部署工具

功能：
  1. 将项目源码部署到 ~/.workbuddy/skills/receiving-registration/
  2. 版本对比：仅当源码版本 > 已部署版本时执行更新
  3. 更新前自动备份到 .backup/<version>/
  4. 仅更新 Skill 文件（SKILL.md / references / scripts），绝不触碰用户数据（~/收货登记/*.xlsx）

使用方式：
  python3 deploy.py              # 检查并部署
  python3 deploy.py --force      # 强制重新部署（忽略版本对比）
  python3 deploy.py --status     # 仅查看版本状态，不执行部署
  python3 deploy.py --backup     # 仅创建备份，不更新

版本号格式：V1.0 / V1.1 / V2.0
"""

import os
import sys
import shutil
import re
import datetime

# ============================================================
# 路径常量
# ============================================================

# 项目源码根目录（本文件所在目录的上级）
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROJECT_SCRIPTS = os.path.join(PROJECT_ROOT, "scripts")
PROJECT_VERSION_FILE = os.path.join(PROJECT_SCRIPTS, "VERSION")
PROJECT_CHANGELOG = os.path.join(PROJECT_SCRIPTS, "CHANGELOG.md")

# 部署目标路径
DEPLOY_TARGET = os.path.expanduser("~/.workbuddy/skills/receiving-registration")
DEPLOY_SCRIPTS = os.path.join(DEPLOY_TARGET, "scripts")
DEPLOY_VERSION_FILE = os.path.join(DEPLOY_SCRIPTS, "VERSION")

# 需要部署的文件清单（相对路径，相对于项目根目录）
DEPLOY_FILES = [
    "SKILL.md",
    "references/field_schema.md",
    "references/company_registry.md",
    "scripts/receiving_excel.py",
    "scripts/VERSION",
    "scripts/CHANGELOG.md",
    "scripts/deploy.py",
]

# 需要部署的目录
DEPLOY_DIRS = [
    "references",
    "scripts",
]


# ============================================================
# 版本解析
# ============================================================

def parse_version(v_str):
    """解析版本号为可比较的元组，支持 V1.0 / 1.0 / V1.0.0 等格式"""
    if not v_str:
        return (0, 0, 0)
    v_str = v_str.strip().upper().lstrip('V')
    parts = v_str.split('.')
    try:
        result = []
        for p in parts[:3]:
            result.append(int(p))
        while len(result) < 3:
            result.append(0)
        return tuple(result)
    except (ValueError, TypeError):
        return (0, 0, 0)


def format_version(v_tuple):
    """格式化版本元组为 V1.0 格式"""
    if len(v_tuple) >= 2:
        return f"V{v_tuple[0]}.{v_tuple[1]}"
    return "V0.0"


def compare_version(v1, v2):
    """比较版本号：v1 < v2 返回 -1，= 返回 0，> 返回 1"""
    a = parse_version(v1)
    b = parse_version(v2)
    if a < b:
        return -1
    elif a > b:
        return 1
    return 0


# ============================================================
# 版本读取
# ============================================================

def get_source_version():
    """读取项目源码版本"""
    if os.path.exists(PROJECT_VERSION_FILE):
        with open(PROJECT_VERSION_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return "V0.0"


def get_deployed_version():
    """读取已部署版本"""
    if os.path.exists(DEPLOY_VERSION_FILE):
        with open(DEPLOY_VERSION_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return None  # 未部署


# ============================================================
# 备份机制
# ============================================================

def create_backup():
    """更新前备份当前已部署的文件"""
    if not os.path.exists(DEPLOY_TARGET):
        return None, "目标目录不存在，无需备份"

    deployed_v = get_deployed_version() or "unknown"
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = os.path.join(DEPLOY_TARGET, ".backup", f"{deployed_v}_{timestamp}")

    os.makedirs(backup_dir, exist_ok=True)

    for rel_path in DEPLOY_FILES:
        src = os.path.join(DEPLOY_TARGET, rel_path)
        if os.path.exists(src):
            dst = os.path.join(backup_dir, rel_path)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)

    return backup_dir, f"备份到 {backup_dir}"


# ============================================================
# 部署逻辑
# ============================================================

def deploy(force=False):
    """执行部署"""
    source_v = get_source_version()
    deployed_v = get_deployed_version()

    print(f"\n{'='*50}")
    print(f"  收货信息登记 Skill 部署工具")
    print(f"{'='*50}")
    print(f"  源码版本:   {source_v}")
    print(f"  已部署版本: {deployed_v or '未部署'}")
    print(f"{'='*50}")

    # 版本对比
    if not force and deployed_v is not None:
        cmp = compare_version(deployed_v, source_v)
        if cmp >= 0:
            print(f"\n  已是最新版本，无需部署")
            return True, "已是最新版本"

    # 创建备份
    if deployed_v is not None and os.path.exists(DEPLOY_TARGET):
        backup_dir, msg = create_backup()
        print(f"  备份: {msg}")
    else:
        print(f"  首次部署，无需备份")

    # 创建目录
    os.makedirs(DEPLOY_TARGET, exist_ok=True)
    os.makedirs(DEPLOY_SCRIPTS, exist_ok=True)
    os.makedirs(os.path.join(DEPLOY_TARGET, "references"), exist_ok=True)

    # 复制文件
    updated = []
    failed = []

    for rel_path in DEPLOY_FILES:
        src = os.path.join(PROJECT_ROOT, rel_path)
        dst = os.path.join(DEPLOY_TARGET, rel_path)

        if not os.path.exists(src):
            failed.append(f"{rel_path} (源文件不存在)")
            continue

        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)
            updated.append(rel_path)
        except Exception as e:
            failed.append(f"{rel_path}: {e}")

    # 输出结果
    print(f"\n  部署完成:")
    print(f"  - 已更新 {len(updated)} 个文件")
    for f in updated:
        print(f"    > {f}")

    if failed:
        print(f"\n  - 失败 {len(failed)} 个:")
        for f in failed:
            print(f"    ! {f}")

    print(f"\n  部署路径: {DEPLOY_TARGET}")
    print(f"  当前版本: {source_v}")
    print(f"  用户数据: ~/收货登记/*.xlsx (未受影响)")
    print(f"{'='*50}\n")

    if failed and not updated:
        return False, f"部署失败: {', '.join(failed[:3])}"
    elif failed:
        return True, f"部署成功（部分失败）: {', '.join(failed[:3])}"
    else:
        return True, f"部署成功: {len(updated)} 个文件"


def show_status():
    """显示版本状态"""
    source_v = get_source_version()
    deployed_v = get_deployed_version()

    print(f"\n  收货信息登记 Skill 版本状态")
    print(f"  {'='*40}")
    print(f"  源码版本:   {source_v}")
    print(f"  已部署版本: {deployed_v or '未部署'}")

    if deployed_v is None:
        print(f"  状态: 未部署")
    else:
        cmp = compare_version(deployed_v, source_v)
        if cmp < 0:
            print(f"  状态: 需要更新 ({deployed_v} -> {source_v})")
        elif cmp > 0:
            print(f"  状态: 已部署版本高于源码（异常）")
        else:
            print(f"  状态: 已是最新")

    print(f"  项目路径:   {PROJECT_ROOT}")
    print(f"  部署路径:   {DEPLOY_TARGET}")
    print(f"  {'='*40}\n")


def backup_only():
    """仅创建备份"""
    backup_dir, msg = create_backup()
    if backup_dir:
        print(f"  备份完成: {msg}")
    else:
        print(f"  {msg}")


# ============================================================
# 命令行入口
# ============================================================

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="收货信息登记 Skill 部署工具")
    parser.add_argument("--force", action="store_true", help="强制重新部署")
    parser.add_argument("--status", action="store_true", help="仅查看版本状态")
    parser.add_argument("--backup", action="store_true", help="仅创建备份")

    args = parser.parse_args()

    if args.status:
        show_status()
    elif args.backup:
        backup_only()
    else:
        success, msg = deploy(force=args.force)
        if not success:
            sys.exit(1)
