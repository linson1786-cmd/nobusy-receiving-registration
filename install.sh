#!/bin/bash
# ============================================================
# 收货信息登记 Skill - 一键安装脚本
# 仓库: https://github.com/linson1786-cmd/nobusy-receiving-registration
# 用法: curl -sL https://raw.githubusercontent.com/linson1786-cmd/nobusy-receiving-registration/main/install.sh | bash
# ============================================================

set -e

REPO="linson1786-cmd/nobusy-receiving-registration"
RAW_BASE="https://raw.githubusercontent.com/${REPO}/main"
DEPLOY_PY_URL="${RAW_BASE}/scripts/receiving-registration/deploy.py"
SKILL_NAME="收货信息登记"
DEPLOY_PATH="$HOME/.workbuddy/skills/receiving-registration"

echo ""
echo "=================================================="
echo "  ${SKILL_NAME} Skill - 一键安装/升级"
echo "=================================================="
echo ""

# 检查 python3
if ! command -v python3 &> /dev/null; then
    echo "  [错误] 未找到 python3，请先安装 Python 3.8+"
    exit 1
fi

# 检查是否已安装
if [ -f "$DEPLOY_PATH/scripts/receiving-registration/VERSION" ]; then
    CURRENT_VER=$(cat "$DEPLOY_PATH/scripts/receiving-registration/VERSION")
    echo "  当前版本: ${CURRENT_VER}"
    echo "  正在检查更新..."
else
    echo "  首次安装"
    echo "  正在下载..."
fi

# 下载 deploy.py 到临时目录
TMP_FILE=$(mktemp /tmp/receiving_deploy_XXXXXX.py)
curl -sL "$DEPLOY_PY_URL" -o "$TMP_FILE"

if [ ! -s "$TMP_FILE" ]; then
    echo "  [错误] 下载 deploy.py 失败"
    rm -f "$TMP_FILE"
    exit 1
fi

# 执行升级（--upgrade 会自动检查版本、下载、备份、部署）
python3 "$TMP_FILE" --upgrade

# 清理
rm -f "$TMP_FILE"

echo ""
echo "=================================================="
echo "  安装/升级完成!"
echo "  安装路径: ${DEPLOY_PATH}"
echo ""
echo "  后续升级方式:"
echo "    1. 对话中说「升级收货登记 Skill」"
echo "    2. 或运行: python3 ~/.workbuddy/skills/receiving-registration/scripts/receiving-registration/deploy.py --upgrade"
echo "=================================================="
echo ""
