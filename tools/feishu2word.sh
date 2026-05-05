#!/bin/sh
SCRIPT_DIR="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"

cd "$PROJECT_ROOT" || exit 1

PYTHON_BIN=""
for CANDIDATE in python3 python; do
    if command -v "$CANDIDATE" >/dev/null 2>&1 && "$CANDIDATE" -c "import docx" >/dev/null 2>&1; then
        PYTHON_BIN="$CANDIDATE"
        break
    fi
done

if [ -z "$PYTHON_BIN" ]; then
    if command -v python3 >/dev/null 2>&1; then
        PYTHON_BIN="python3"
    elif command -v python >/dev/null 2>&1; then
        PYTHON_BIN="python"
    else
        echo "Python is not installed or not in PATH." >&2
        exit 1
    fi
fi

exec "$PYTHON_BIN" "src/cli/feishu2word.py" "$@"
