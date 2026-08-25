#!/usr/bin/env python3
"""将旧版 users.json / download_stats.jsonl 导入 SQLite。

默认只导入、不删除旧文件；迁移成功后会在同目录留下带时间戳的 .bak 备份。
脚本可重复执行，已迁移过的来源会被跳过；确需重新导入时使用 --force。
"""

import argparse
import json
import os
import sys


PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from src.core.config_loader import config as loaded_config
from src.core.sqlite_store import migrate_legacy_data


def parse_args():
    parser = argparse.ArgumentParser(description='迁移 users.json 和 download_stats.jsonl 到 SQLite')
    parser.add_argument('--workspace', help='工作区目录，默认使用配置 workspace.dir')
    parser.add_argument('--log-dir', help='日志目录，默认使用配置 log.dir')
    parser.add_argument('--db', help='自定义 SQLite 文件路径，默认 <workspace>/<log.dir>/feishu_docget.sqlite3')
    parser.add_argument('--dry-run', action='store_true', help='只统计可迁移数据，不创建或修改数据库')
    parser.add_argument('--no-backup', action='store_true', help='不额外复制旧 JSON/JSONL 备份')
    parser.add_argument('--force', action='store_true', help='忽略已迁移标记并覆盖同名记录，谨慎使用')
    return parser.parse_args()


def main():
    args = parse_args()
    migration_config = dict(loaded_config)
    workspace = os.path.abspath(args.workspace or migration_config.get('workspace.dir', '.'))
    migration_config['workspace.dir'] = workspace
    if args.log_dir:
        migration_config['log.dir'] = args.log_dir
    result = migrate_legacy_data(
        workspace,
        migration_config,
        db_path=os.path.abspath(args.db) if args.db else None,
        backup=not args.no_backup,
        dry_run=args.dry_run,
        force=args.force,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
