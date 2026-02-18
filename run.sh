#!/bin/bash
# ============================================================
#  当直くん - ワンクリック実行スクリプト (macOS用)
#
#  使い方: ターミナルに以下をコピー＆ペーストするだけ
#    bash run.sh
#
#  または、リポジトリのクローンから一発実行:
#    git clone https://github.com/tomitayus/Tochoku-kun.git && cd Tochoku-kun && bash run.sh
# ============================================================
set -euo pipefail

# --- カラー定義 ---
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[0;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
BOLD='\033[1m'
NC='\033[0m' # No Color

# --- スクリプトのディレクトリに移動 ---
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# ============================================================
#  バナー表示
# ============================================================
echo ""
echo -e "${CYAN}${BOLD}============================================================${NC}"
echo -e "${CYAN}${BOLD}  当直くん - 医師当直スケジュール自動生成ツール${NC}"
echo -e "${CYAN}${BOLD}============================================================${NC}"
echo ""
echo -e "  ${BOLD}セットアップから実行まで全自動で行います。${NC}"
echo -e "  Excelファイルの選択ダイアログが表示されるので、"
echo -e "  入力ファイルを選択してください。"
echo ""

# ============================================================
#  Python 3 チェック
# ============================================================
echo -e "${BLUE}[1/4]${NC} Python 3 をチェック中..."

PYTHON_CMD=""
if command -v python3 &>/dev/null; then
    PYTHON_CMD="python3"
elif command -v python &>/dev/null; then
    PY_VER=$(python --version 2>&1 | grep -oE '[0-9]+\.[0-9]+')
    PY_MAJOR=$(echo "$PY_VER" | cut -d. -f1)
    if [ "$PY_MAJOR" = "3" ]; then
        PYTHON_CMD="python"
    fi
fi

if [ -z "$PYTHON_CMD" ]; then
    echo ""
    echo -e "${RED}エラー: Python 3 が見つかりません。${NC}"
    echo ""
    echo "  macOS に Python 3 をインストールする方法:"
    echo ""
    echo "  方法1: Homebrew (推奨)"
    echo "    /bin/bash -c \"\$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\""
    echo "    brew install python"
    echo ""
    echo "  方法2: python.org から直接ダウンロード"
    echo "    https://www.python.org/downloads/"
    echo ""
    echo "  インストール後、もう一度このスクリプトを実行してください。"
    exit 1
fi

PY_VERSION=$($PYTHON_CMD --version 2>&1)
echo -e "  ${GREEN}OK${NC}: $PY_VERSION"

# バージョン 3.10 以上かチェック
PY_MINOR=$($PYTHON_CMD -c "import sys; print(sys.version_info.minor)")
PY_MAJOR=$($PYTHON_CMD -c "import sys; print(sys.version_info.major)")
if [ "$PY_MAJOR" -lt 3 ] || { [ "$PY_MAJOR" -eq 3 ] && [ "$PY_MINOR" -lt 10 ]; }; then
    echo -e "${YELLOW}警告: Python 3.10 以上を推奨します (現在: $PY_VERSION)${NC}"
fi

# ============================================================
#  仮想環境のセットアップ
# ============================================================
echo -e "${BLUE}[2/4]${NC} 仮想環境をセットアップ中..."

VENV_DIR="$SCRIPT_DIR/.venv"

if [ ! -d "$VENV_DIR" ]; then
    echo "  仮想環境を作成しています..."
    $PYTHON_CMD -m venv "$VENV_DIR"
    echo -e "  ${GREEN}OK${NC}: .venv を作成しました"
else
    echo -e "  ${GREEN}OK${NC}: 既存の .venv を使用します"
fi

# 仮想環境をアクティベート
source "$VENV_DIR/bin/activate"

# ============================================================
#  依存パッケージのインストール
# ============================================================
echo -e "${BLUE}[3/4]${NC} 依存パッケージをインストール中..."

# 必要なパッケージがすでにインストール済みかチェック
NEED_INSTALL=false
for pkg in pandas numpy openpyxl; do
    if ! python -c "import $pkg" &>/dev/null; then
        NEED_INSTALL=true
        break
    fi
done

if [ "$NEED_INSTALL" = true ]; then
    echo "  パッケージをインストールしています（初回のみ）..."
    pip install --quiet --upgrade pip
    pip install --quiet -r "$SCRIPT_DIR/requirements.txt"
    echo -e "  ${GREEN}OK${NC}: パッケージをインストールしました"
else
    echo -e "  ${GREEN}OK${NC}: パッケージは全てインストール済みです"
fi

# ============================================================
#  当直表生成の実行
# ============================================================
echo -e "${BLUE}[4/4]${NC} 当直表を生成します..."
echo ""
echo -e "${YELLOW}============================================================${NC}"
echo -e "${YELLOW}  Excelファイルの選択ダイアログが表示されます。${NC}"
echo -e "${YELLOW}  入力ファイル（sheet1〜4を含むExcel）を選択してください。${NC}"
echo -e "${YELLOW}============================================================${NC}"
echo ""

# main.py を実行（osascript でファイル選択ダイアログが自動表示される）
set +e
python "$SCRIPT_DIR/main.py"
EXIT_CODE=$?
set -e

# ============================================================
#  完了メッセージ
# ============================================================
echo ""
if [ $EXIT_CODE -eq 0 ]; then
    echo -e "${GREEN}${BOLD}============================================================${NC}"
    echo -e "${GREEN}${BOLD}  完了しました!${NC}"
    echo -e "${GREEN}${BOLD}============================================================${NC}"
    echo ""
    echo -e "  出力ファイルは ${BOLD}~/Downloads/${NC} フォルダに保存されています。"
    echo ""

    # Finder で出力フォルダを開く（macOS のみ）
    if command -v open &>/dev/null; then
        DOWNLOADS_DIR="$HOME/Downloads"
        # 最新の出力ファイルを見つけて開く
        LATEST_FILE=$(ls -t "$DOWNLOADS_DIR"/*_v*.xlsx 2>/dev/null | head -1)
        if [ -n "$LATEST_FILE" ]; then
            echo -e "  出力ファイル: ${CYAN}$(basename "$LATEST_FILE")${NC}"
            echo ""
            read -p "  Finder で出力フォルダを開きますか? [Y/n] " -n 1 -r
            echo ""
            if [[ ! $REPLY =~ ^[Nn]$ ]]; then
                open -R "$LATEST_FILE"
            fi
        fi
    fi
else
    echo -e "${RED}${BOLD}============================================================${NC}"
    echo -e "${RED}${BOLD}  エラーが発生しました (終了コード: $EXIT_CODE)${NC}"
    echo -e "${RED}${BOLD}============================================================${NC}"
    echo ""
    echo "  よくある原因:"
    echo "    - Excelファイルが正しい形式（sheet1〜4）でない"
    echo "    - ファイル選択がキャンセルされた"
    echo ""
    echo "  詳しくは上のエラーメッセージを確認してください。"
fi

echo ""
