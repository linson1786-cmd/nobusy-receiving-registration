#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
收货信息登记 Skill 部署工具

功能：
  1. 将项目源码部署到 ~/.workbuddy/skills/receiving-registration/
  2. 版本对比：仅当源码版本 > 已部署版本时执行更新
  3. 更新前自动备份到 .backup/<version>/
  4. 仅更新 Skill 文件（SKILL.md / references / scripts），绝不触碰用户数据（~/收货登记/*.xlsx）
  5. --upgrade：从 GitHub 自动拉取最新版本并部署

使用方式：
  python3 deploy.py              # 检查并部署（从本地源码）
  python3 deploy.py --force      # 强制重新部署（忽略版本对比）
  python3 deploy.py --status     # 仅查看版本状态，不执行部署
  python3 deploy.py --backup     # 仅创建备份，不更新
  python3 deploy.py --upgrade    # 从 GitHub 拉取最新版本并自动部署
  python3 deploy.py --check-update  # 仅检查是否有新版本可用

版本号格式：V1.0 / V1.1 / V2.0
"""

import os
import sys
import shutil
import re
import datetime
import json
import urllib.request
import urllib.error

# ============================================================
# 路径常量
# ============================================================

# 项目源码根目录（本文件位于 scripts/receiving-registration/）
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
PROJECT_SCRIPTS = os.path.join(PROJECT_ROOT, "scripts", "receiving-registration")
PROJECT_VERSION_FILE = os.path.join(PROJECT_SCRIPTS, "VERSION")
PROJECT_CHANGELOG = os.path.join(PROJECT_SCRIPTS, "CHANGELOG.md")

# 部署目标路径
DEPLOY_TARGET = os.path.expanduser("~/.workbuddy/skills/receiving-registration")
DEPLOY_SCRIPTS = os.path.join(DEPLOY_TARGET, "scripts", "receiving-registration")
DEPLOY_VERSION_FILE = os.path.join(DEPLOY_SCRIPTS, "VERSION")

# 需要部署的文件清单（相对路径，相对于项目根目录/交付包根目录）
DEPLOY_FILES = [
    "SKILL.md",
    "references/field_schema.md",
    "references/company_registry.md",
    "scripts/receiving-registration/receiving_excel.py",
    "scripts/receiving-registration/VERSION",
    "scripts/receiving-registration/CHANGELOG.md",
    "scripts/receiving-registration/deploy.py",
]

# 需要部署的目录
DEPLOY_DIRS = [
    "references",
    "scripts/receiving-registration",
]

# ============================================================
# 远程仓库配置（用于 --upgrade）
# ============================================================

GITHUB_REPO = "https://github.com/linson1786-cmd/nobusy-receiving-registration.git"
# 交付包在仓库中的路径（相对于仓库根目录）
REPO_DELIVERY_PATH = "交付包"


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
# 本地部署逻辑（从本地源码到部署目录）
# ============================================================

def deploy(force=False):
    """执行部署（从本地 PROJECT_ROOT 复制到 DEPLOY_TARGET）"""
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
# 远程升级功能（从 GitHub 拉取）
# ============================================================

def _run_git(args, cwd=None):
    """执行 git 命令，返回 (returncode, stdout, stderr)"""
    import subprocess
    try:
        result = subprocess.run(
            ["git"] + args,
            cwd=cwd,
            capture_output=True,
            text=True,
            timeout=60
        )
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except FileNotFoundError:
        return -1, "", "git 命令未找到，请先安装 git"
    except subprocess.TimeoutExpired:
        return -1, "", "git 命令超时"


def check_remote_update():
    """检查 GitHub 远程仓库是否有新版本可用"""
    deployed_v = get_deployed_version()

    print(f"\n  检查远程更新...")
    print(f"  已部署版本: {deployed_v or '未部署'}")
    print(f"  仓库地址:   {GITHUB_REPO}")
    print(f"  {'='*40}")

    latest_remote = None

    # 方法1：尝试通过 git ls-remote 获取最新 tag
    rc, stdout, stderr = _run_git(["ls-remote", "--tags", "--sort=-v:refname", GITHUB_REPO])
    if rc == 0 and stdout:
        for line in stdout.split('\n'):
            if not line:
                continue
            parts = line.split('\t')
            if len(parts) >= 2:
                ref = parts[1]
                tag_name = ref.replace('refs/tags/', '').lstrip('vV')
                try:
                    parse_version(tag_name)
                    if not latest_remote or compare_version(tag_name, latest_remote) > 0:
                        latest_remote = tag_name
                except Exception:
                    continue

    # 方法2（备用）：通过 GitHub API 获取 tags
    if not latest_remote:
        api_url = GITHUB_REPO.replace('.git', '') + '/tags?per_page=5'
        print(f"  尝试 GitHub API: {api_url}")
        try:
            req = urllib.request.Request(api_url, headers={"User-Agent": "deploy-py/1.0"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                tags_data = json.loads(resp.read().decode('utf-8'))
                for tag_info in tags_data:
                    tag_name = tag_info.get("name", "").lstrip('vV')
                    try:
                        parse_version(tag_name)
                        if not latest_remote or compare_version(tag_name, latest_remote) > 0:
                            latest_remote = tag_name
                    except Exception:
                        continue
        except urllib.error.HTTPError as e:
            if e.code == 404:
                print(f"  API 返回 404: 仓库可能不存在或为私有")
            elif e.code == 403:
                print(f"  API 速率限制或认证问题")
            else:
                print(f"  API 错误: HTTP {e.code}")
        except Exception as e:
            print(f"  API 请求失败: {e}")

    # 输出结果
    if latest_remote:
        remote_v_str = f"V{latest_remote}"
        print(f"  最新远程版本: {remote_v_str}")

        if deployed_v is None:
            print(f"  状态: 尚未部署，可安装 {remote_v_str}")
            return True, f"可安装: {remote_v_str}", remote_v_str

        cmp = compare_version(deployed_v, remote_v_str)
        if cmp < 0:
            print(f"  状态: 有新版本可用! ({deployed_v} -> {remote_v_str})")
            return True, f"有新版本: {deployed_v} -> {remote_v_str}", remote_v_str
        elif cmp > 0:
            print(f"  状态: 本地版本高于远程（可能是开发版）")
            return False, "本地已是最新(开发版)", None
        else:
            print(f"  状态: 已是最新版本")
            return False, "已是最新版本", None

    print(f"  无法获取远程版本信息（请检查网络连接和仓库权限）")
    return False, "无法获取远程版本信息", None


def _do_deploy_from(source_root, target_label=""):
    """
    从指定源目录部署文件到 DEPLOY_TARGET 的内部实现。
    source_root: SKILL.md 所在的根目录
    target_label: 日志中显示的来源标识
    """
    # 读取新版本号
    src_ver_file = os.path.join(source_root, "scripts", "receiving-registration", "VERSION")
    new_version = "unknown"
    if os.path.exists(src_ver_file):
        with open(src_ver_file, 'r', encoding='utf-8') as f:
            new_version = f.read().strip()

    # 备份
    deployed_v = get_deployed_version()
    if deployed_v and os.path.exists(DEPLOY_TARGET):
        backup_dir, backup_msg = create_backup()
        print(f"  备份: {backup_msg}")
    else:
        print(f"  首次安装，无需备份")

    # 创建目录
    os.makedirs(DEPLOY_TARGET, exist_ok=True)
    os.makedirs(DEPLOY_SCRIPTS, exist_ok=True)
    os.makedirs(os.path.join(DEPLOY_TARGET, "references"), exist_ok=True)

    # 复制文件
    updated = []
    failed = []

    for rel_path in DEPLOY_FILES:
        src = os.path.join(source_root, rel_path)
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
    action = f"升级({target_label})" if target_label else "部署"
    print(f"\n  {action}完成:")
    print(f"  - 已更新 {len(updated)} 个文件")
    for f in updated:
        print(f"    > {f}")

    if failed:
        print(f"\n  - 失败 {len(failed)} 个:")
        for f in failed:
            print(f"    ! {f}")

    print(f"\n  部署路径: {DEPLOY_TARGET}")
    print(f"  当前版本: {new_version}")
    print(f"  用户数据: ~/收货登记/*.xlsx (未受影响)")
    print(f"{'='*50}\n")

    old_ver = deployed_v or "未安装"

    if failed and not updated:
        return False, f"{action}失败: {', '.join(failed[:3])}"
    elif failed:
        return True, f"{action}成功（部分失败）: {', '.join(failed[:3])}"
    else:
        return True, f"{action}完成: {old_ver} -> {new_version}"


def upgrade(force=False):
    """从 GitHub 自动拉取最新版本并部署"""
    print(f"\n{'='*50}")
    print(f"  收货信息登记 Skill - 在线升级")
    print(f"{'='*50}")

    # 第一步：检查远程是否有新版本
    has_update, update_msg, remote_version = check_remote_update()

    if not has_update and not force:
        print(f"\n  无需更新: {update_msg}")
        return True, update_msg

    if force and not remote_version:
        print(f"\n  强制模式：但无法确定远程版本")
        return False, "强制升级失败：无法获取远程版本"

    # 第二步：确认目标版本
    target_version = f"V{remote_version}" if remote_version else "latest"
    print(f"\n  目标版本: {target_version}")

    # 第三步：创建临时目录，克隆仓库
    import tempfile
    tmp_dir = tempfile.mkdtemp(prefix="receiving_upgrade_")
    print(f"  下载中... (临时目录: {tmp_dir})")

    try:
        # 克隆仓库（浅克隆，只取最新）
        rc, stdout, stderr = _run_git(
            ["clone", "--depth", "1", GITHUB_REPO, tmp_dir]
        )
        clone_ok = (rc == 0)

        # git 失败时，尝试用 HTTP 下载 zip
        if not clone_ok:
            print(f"  git 克隆失败: {stderr}")
            print(f"  尝试通过 GitHub API 下载...")
            zip_url = GITHUB_REPO.replace('.git', '') + '/archive/refs/heads/main.zip'
            try:
                req = urllib.request.Request(zip_url, headers={"User-Agent": "deploy-py/1.0"})
                with urllib.request.urlopen(req, timeout=60) as resp:
                    zip_data = resp.read()
                zip_path = os.path.join(tmp_dir, "repo.zip")
                # tmp_dir 本身就是临时目录，我们在其父目录下创建真正的解压目标
                import zipfile
                parent_tmp = os.path.dirname(tmp_dir)
                extract_dir = os.path.join(parent_tmp, "receiving_extract")
                os.makedirs(extract_dir, exist_ok=True)

                with open(zip_path, 'wb') as zf:
                    zf.write(zip_data)

                with zipfile.ZipFile(zip_path, 'r') as zf:
                    zf.extractall(extract_dir)

                # zip 解压后通常有一个根子目录（如 repo-main），找出来
                extracted_items = os.listdir(extract_dir)
                if len(extracted_items) == 1:
                    actual_root = os.path.join(extract_dir, extracted_items[0])
                    if os.path.isdir(actual_root):
                        # 把内容移到 tmp_dir
                        for item in os.listdir(actual_root):
                            shutil.move(os.path.join(actual_root, item), tmp_dir)
                        shutil.rmtree(actual_root)
                        shutil.rmtree(extract_dir)
                        clone_ok = True
                        print(f"  ZIP 下载成功")
                    else:
                        print(f"  ZIP 结构异常")
                if not clone_ok:
                    # 尝试用 extract_dir 直接作为根目录
                    if os.path.exists(os.path.join(extract_dir, "SKILL.md")):
                        # 把 extract_dir 的内容复制到 tmp_dir
                        for item in os.listdir(extract_dir):
                            src = os.path.join(extract_dir, item)
                            dst = os.path.join(tmp_dir, item)
                            if os.path.isdir(src):
                                shutil.copytree(src, dst)
                            else:
                                shutil.copy2(src, dst)
                        shutil.rmtree(extract_dir)
                        clone_ok = True
                        print(f"  ZIP 下载成功（直接结构）")
            except Exception as e:
                print(f"  ZIP 下载也失败: {e}")

        if not clone_ok:
            return False, "无法从 GitHub 获取代码（git 和 HTTP 均失败）"

        print(f"  下载成功")

        # 第四步：定位交付包或仓库根目录作为文件来源
        clone_delivery = os.path.join(tmp_dir, REPO_DELIVERY_PATH)
        source_root_for_deploy = None

        # 优先找交付包/<version>/ 目录
        if os.path.isdir(clone_delivery):
            version_dirs = []
            for d in os.listdir(clone_delivery):
                full_path = os.path.join(clone_delivery, d)
                if os.path.isdir(full_path) and (d.startswith("V") or d.startswith("v")):
                    version_dirs.append((d, parse_version(d)))
            version_dirs.sort(key=lambda x: x[1], reverse=True)

            if version_dirs:
                candidate = os.path.join(clone_delivery, version_dirs[0][0])
                if os.path.exists(os.path.join(candidate, "SKILL.md")):
                    source_root_for_deploy = candidate
                    print(f"  使用交付包: {version_dirs[0][0]}")

        # 回退：用仓库根目录
        if not source_root_for_deploy:
            if os.path.exists(os.path.join(tmp_dir, "SKILL.md")):
                source_root_for_deploy = tmp_dir
                print(f"  使用仓库根目录作为源")
            else:
                print(f"  错误: 找不到 SKILL.md，仓库结构可能已变化")
                return False, "找不到有效的 Skill 文件"

        # 第五步：执行部署
        return _do_deploy_from(source_root_for_deploy, target_label=target_version)

    finally:
        # 清理临时目录
        try:
            shutil.rmtree(tmp_dir)
        except Exception:
            pass


# ============================================================
# 命令行入口
# ============================================================

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="收货信息登记 Skill 部署工具",
        epilog=(
            "示例:\n"
            "  python3 deploy.py                 # 从本地源码检查并部署\n"
            "  python3 deploy.py --upgrade        # 从 GitHub 升级到最新版\n"
            "  python3 deploy.py --check-update   # 检查是否有新版本\n"
            "  python3 deploy.py --status         # 查看版本状态\n"
            "  python3 deploy.py --force          # 强制重新部署\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--force", action="store_true", help="强制重新部署（忽略版本对比）")
    parser.add_argument("--status", action="store_true", help="仅查看版本状态，不执行部署")
    parser.add_argument("--backup", action="store_true", help="仅创建备份，不更新")
    parser.add_argument("--upgrade", action="store_true", help="从 GitHub 拉取最新版本并自动部署")
    parser.add_argument("--check-update", action="store_true", help="检查远程仓库是否有新版本可用")

    args = parser.parse_args()

    if args.upgrade:
        success, msg = upgrade(force=args.force)
        if not success:
            sys.exit(1)
    elif args.check_update:
        check_remote_update()
    elif args.status:
        show_status()
    elif args.backup:
        backup_only()
    else:
        success, msg = deploy(force=args.force)
        if not success:
            sys.exit(1)
