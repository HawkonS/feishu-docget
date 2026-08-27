"""SQLite 持久化层。

统计和用户数据都放在同一个 SQLite 文件中，避免 JSON/JSONL 全量扫描。
本模块只依赖 Python 标准库 sqlite3；旧数据迁移逻辑也放在这里，供应用
启动时的一次性兼容迁移和 tools/migrate_json_to_sqlite.py 复用。
"""

import hashlib
import json
import os
import shutil
import sqlite3
import threading
import time
from datetime import datetime, timedelta
from urllib.parse import urlsplit


DATABASE_FILENAME = 'feishu_docget.sqlite3'
SCHEMA_VERSION = '4'
_schema_lock = threading.RLock()
_initialized_paths = set()


def database_path(base_dir, config):
    """返回数据库路径，默认与日志/用户库放在同一受保护目录。"""
    workspace = os.path.abspath(base_dir or config.get('workspace.dir', '.'))
    log_dir = str(config.get('log.dir', 'logs') or 'logs')
    path = os.path.join(workspace, log_dir, DATABASE_FILENAME)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    return path


def _secure_file(path):
    try:
        os.chmod(os.path.realpath(path), 0o600)
    except OSError:
        pass


def connect(base_dir, config, db_path=None):
    """创建一个短生命周期连接，并设置适合单机 Flask 的 SQLite 参数。"""
    path = db_path or database_path(base_dir, config)
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    connection = sqlite3.connect(path, timeout=30)
    connection.row_factory = sqlite3.Row
    connection.execute('PRAGMA busy_timeout=30000')
    connection.execute('PRAGMA foreign_keys=ON')
    connection.execute('PRAGMA journal_mode=WAL')
    connection.execute('PRAGMA synchronous=NORMAL')
    _secure_file(path)
    _secure_file(path + '-wal')
    _secure_file(path + '-shm')
    return connection


def initialize_database(base_dir, config, db_path=None):
    """创建表、索引和迁移元数据；重复执行安全。"""
    path = db_path or database_path(base_dir, config)
    with _schema_lock:
        if path in _initialized_paths and os.path.isfile(path):
            return
        connection = connect(base_dir, config, db_path=path)
        try:
            connection.executescript(
                """
            CREATE TABLE IF NOT EXISTS schema_meta (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS users (
                open_id TEXT PRIMARY KEY,
                union_id TEXT NOT NULL DEFAULT '',
                user_id TEXT NOT NULL DEFAULT '',
                name TEXT NOT NULL DEFAULT '',
                avatar TEXT NOT NULL DEFAULT '',
                disabled INTEGER NOT NULL DEFAULT 0,
                is_admin INTEGER NOT NULL DEFAULT 0,
                is_system_admin INTEGER NOT NULL DEFAULT 0,
                access_token TEXT NOT NULL DEFAULT '',
                refresh_token TEXT NOT NULL DEFAULT '',
                token_expire_at INTEGER NOT NULL DEFAULT 0,
                refresh_token_expire_at INTEGER NOT NULL DEFAULT 0,
                scope TEXT NOT NULL DEFAULT '',
                token_invalid INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL DEFAULT '',
                last_login_at TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_users_last_login_at ON users(last_login_at DESC);
            CREATE INDEX IF NOT EXISTS idx_users_name ON users(name COLLATE NOCASE);
            CREATE INDEX IF NOT EXISTS idx_users_disabled ON users(disabled);
            CREATE INDEX IF NOT EXISTS idx_users_is_admin ON users(is_admin);

            CREATE TABLE IF NOT EXISTS download_stats (
                row_id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id TEXT,
                legacy_key TEXT UNIQUE,
                status TEXT NOT NULL DEFAULT '',
                ts INTEGER NOT NULL DEFAULT 0,
                time TEXT NOT NULL DEFAULT '',
                url TEXT NOT NULL DEFAULT '',
                path TEXT NOT NULL DEFAULT '',
                title TEXT NOT NULL DEFAULT '',
                ip TEXT NOT NULL DEFAULT '',
                user TEXT NOT NULL DEFAULT ''
            );
            CREATE UNIQUE INDEX IF NOT EXISTS idx_stats_task_id
                ON download_stats(task_id) WHERE task_id IS NOT NULL;
            CREATE INDEX IF NOT EXISTS idx_stats_ts ON download_stats(ts DESC);
            CREATE INDEX IF NOT EXISTS idx_stats_time ON download_stats(time);
            CREATE INDEX IF NOT EXISTS idx_stats_title ON download_stats(title COLLATE NOCASE);
            CREATE INDEX IF NOT EXISTS idx_stats_user ON download_stats(user COLLATE NOCASE);
            CREATE INDEX IF NOT EXISTS idx_stats_ip ON download_stats(ip COLLATE NOCASE);
                """
            )
            # Role metadata was introduced after the original schema. Migrate
            # the former single-binding field into the system-admin role.
            user_columns = {
                row['name'] for row in connection.execute('PRAGMA table_info(users)').fetchall()
            }
            if 'is_system_admin' not in user_columns:
                connection.execute(
                    'ALTER TABLE users ADD COLUMN is_system_admin INTEGER NOT NULL DEFAULT 0'
                )
            if 'system_admin_bound' in user_columns:
                connection.execute(
                    'UPDATE users SET is_system_admin=1 WHERE system_admin_bound=1'
                )
                # Clear the retired field so a later restart cannot restore a
                # system-admin role that has since been removed.
                connection.execute('UPDATE users SET system_admin_bound=0')
            connection.execute('DROP INDEX IF EXISTS idx_users_single_system_admin_bound')
            _set_meta(connection, 'schema_version', SCHEMA_VERSION)
            connection.commit()
            _initialized_paths.add(path)
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()


def _set_meta(connection, key, value):
    connection.execute(
        "INSERT INTO schema_meta(key, value) VALUES(?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, str(value)),
    )


def get_meta(base_dir, config, key, db_path=None):
    initialize_database(base_dir, config, db_path=db_path)
    connection = connect(base_dir, config, db_path=db_path)
    try:
        row = connection.execute('SELECT value FROM schema_meta WHERE key=?', (key,)).fetchone()
        return row['value'] if row else None
    finally:
        connection.close()


def _get_meta_from_connection(connection, key):
    row = connection.execute('SELECT value FROM schema_meta WHERE key=?', (key,)).fetchone()
    return row['value'] if row else None


def _normalise_user(record, open_id=None):
    record = dict(record or {})
    if 'is_system_admin' not in record:
        record['is_system_admin'] = record.get('system_admin_bound', False)
    open_id = str(record.get('open_id') or open_id or '').strip()
    if not open_id:
        return None
    integer_fields = ('token_expire_at', 'refresh_token_expire_at')
    for field in integer_fields:
        try:
            record[field] = int(record.get(field) or 0)
        except (TypeError, ValueError):
            record[field] = 0
    for field in ('disabled', 'is_admin', 'is_system_admin', 'token_invalid'):
        record[field] = 1 if record.get(field) else 0
    values = {'open_id': open_id}
    for field in (
        'union_id', 'user_id', 'name', 'avatar', 'access_token', 'refresh_token',
        'scope', 'created_at', 'last_login_at',
    ):
        values[field] = str(record.get(field) or '')
    for field in integer_fields + ('disabled', 'is_admin', 'is_system_admin', 'token_invalid'):
        values[field] = record[field]
    return values


def get_user(base_dir, config, open_id):
    if not open_id:
        return None
    initialize_database(base_dir, config)
    connection = connect(base_dir, config)
    try:
        row = connection.execute('SELECT * FROM users WHERE open_id=?', (str(open_id),)).fetchone()
        return dict(row) if row else None
    finally:
        connection.close()


def upsert_user(base_dir, config, record):
    values = _normalise_user(record)
    if not values:
        return False
    initialize_database(base_dir, config)
    connection = connect(base_dir, config)
    try:
        connection.execute(
            """
            INSERT INTO users(
                open_id, union_id, user_id, name, avatar, disabled, is_admin, is_system_admin,
                access_token, refresh_token, token_expire_at,
                refresh_token_expire_at, scope, token_invalid, created_at, last_login_at
            ) VALUES(
                :open_id, :union_id, :user_id, :name, :avatar, :disabled, :is_admin, :is_system_admin,
                :access_token, :refresh_token, :token_expire_at,
                :refresh_token_expire_at, :scope, :token_invalid, :created_at, :last_login_at
            )
            ON CONFLICT(open_id) DO UPDATE SET
                union_id=excluded.union_id,
                user_id=excluded.user_id,
                name=excluded.name,
                avatar=excluded.avatar,
                disabled=excluded.disabled,
                is_admin=excluded.is_admin,
                is_system_admin=excluded.is_system_admin,
                access_token=excluded.access_token,
                refresh_token=excluded.refresh_token,
                token_expire_at=excluded.token_expire_at,
                refresh_token_expire_at=excluded.refresh_token_expire_at,
                scope=excluded.scope,
                token_invalid=excluded.token_invalid,
                created_at=excluded.created_at,
                last_login_at=excluded.last_login_at
            """,
            values,
        )
        connection.commit()
        return True
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def update_user_fields(base_dir, config, open_id, **fields):
    allowed = {
        'disabled', 'is_admin', 'is_system_admin', 'access_token', 'refresh_token', 'token_expire_at',
        'refresh_token_expire_at', 'scope', 'token_invalid', 'name', 'avatar',
        'last_login_at', 'union_id', 'user_id',
    }
    fields = {key: value for key, value in fields.items() if key in allowed}
    if not open_id or not fields:
        return False
    for field in ('disabled', 'is_admin', 'is_system_admin', 'token_invalid'):
        if field in fields:
            fields[field] = 1 if fields[field] else 0
    for field in ('token_expire_at', 'refresh_token_expire_at'):
        if field in fields:
            try:
                fields[field] = int(fields[field] or 0)
            except (TypeError, ValueError):
                fields[field] = 0
    initialize_database(base_dir, config)
    connection = connect(base_dir, config)
    try:
        assignments = ', '.join(f'{key}=?' for key in fields)
        cursor = connection.execute(
            f'UPDATE users SET {assignments} WHERE open_id=?',
            tuple(fields.values()) + (str(open_id),),
        )
        connection.commit()
        return cursor.rowcount > 0
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def set_user_role(base_dir, config, open_id, role):
    """Atomically set one of the mutually exclusive user roles."""
    open_id = str(open_id or '').strip()
    role = str(role or '').strip()
    role_flags = {
        'user': (0, 0),
        'operator_admin': (1, 0),
        'system_admin': (0, 1),
    }
    if not open_id or open_id == '__system_admin__' or role not in role_flags:
        return False
    initialize_database(base_dir, config)
    connection = connect(base_dir, config)
    try:
        target = connection.execute(
            'SELECT disabled FROM users WHERE open_id=?', (open_id,)
        ).fetchone()
        if not target or (role == 'system_admin' and target['disabled']):
            return False
        is_admin, is_system_admin = role_flags[role]
        cursor = connection.execute(
            'UPDATE users SET is_admin=?, is_system_admin=? WHERE open_id=?',
            (is_admin, is_system_admin, open_id),
        )
        connection.commit()
        return cursor.rowcount > 0
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def set_system_admin(base_dir, config, open_id, is_system_admin):
    """Set a Feishu user's system-admin role; multiple users may hold it."""
    return set_user_role(
        base_dir,
        config,
        open_id,
        'system_admin' if is_system_admin else 'user',
    )


def list_users(base_dir, config, page=None, page_size=None, query=''):
    initialize_database(base_dir, config)
    connection = connect(base_dir, config)
    try:
        where = []
        params = []
        query = str(query or '').strip()
        if query:
            where.append('(name LIKE ? COLLATE NOCASE OR open_id LIKE ? COLLATE NOCASE)')
            like = f'%{query}%'
            params.extend((like, like))
        where_sql = (' WHERE ' + ' AND '.join(where)) if where else ''
        total = connection.execute(f'SELECT COUNT(*) AS count FROM users{where_sql}', params).fetchone()['count']
        page_items_sql = (
            'SELECT open_id, union_id, user_id, name, avatar, disabled, is_admin, is_system_admin, '
            'token_expire_at, refresh_token_expire_at, scope, token_invalid, created_at, last_login_at '
            f'FROM users{where_sql} '
            'ORDER BY CASE WHEN open_id = \'__system_admin__\' THEN 0 '
            'WHEN is_system_admin = 1 THEN 1 WHEN is_admin = 1 THEN 2 ELSE 3 END, '
            'last_login_at DESC, open_id ASC'
        )
        page_params = list(params)
        if page is None and page_size is None:
            page = None
        else:
            try:
                page = max(int(page or 1), 1)
            except (TypeError, ValueError):
                page = 1
        if page is not None:
            try:
                page_size = min(max(int(page_size or 20), 1), 100)
            except (TypeError, ValueError):
                page_size = 20
            total_pages = (total + page_size - 1) // page_size if total else 0
            if total_pages:
                page = min(page, total_pages)
            else:
                page = 1
            page_params.extend((page_size, (page - 1) * page_size))
            page_items_sql += ' LIMIT ? OFFSET ?'
        rows = [dict(row) for row in connection.execute(page_items_sql, page_params).fetchall()]
        for row in rows:
            row['disabled'] = bool(row.get('disabled'))
            row['is_admin'] = bool(row.get('is_admin'))
            row['is_system_admin'] = bool(row.get('is_system_admin'))
            row['is_system'] = row.get('open_id') == '__system_admin__'
            row['is_system_admin'] = row['is_system'] or row['is_system_admin']
            row['is_operator_admin'] = row['is_admin'] and not row['is_system_admin']
            row['token_invalid'] = bool(row.get('token_invalid'))
            if not row.get('created_at') and row.get('last_login_at'):
                row['created_at'] = row['last_login_at']
        if page is None:
            return {'items': rows, 'total': total}
        return {
            'items': rows,
            'total': total,
            'page': page,
            'page_size': page_size,
            'total_pages': total_pages,
            'has_more': page * page_size < total,
        }
    finally:
        connection.close()


def _stat_values(entry):
    values = dict(entry or {})
    values['task_id'] = str(values.get('id') or '').strip() or None
    values['status'] = str(values.get('status') or '')
    try:
        values['ts'] = int(values.get('ts') or 0)
    except (TypeError, ValueError):
        values['ts'] = 0
    for field in ('time', 'url', 'path', 'title', 'ip', 'user'):
        values[field] = str(values.get(field) or '')
    values['url'] = _mask_url(values['url'])
    values['ip'] = _mask_ip(values['ip'])
    return values


def _mask_ip(ip):
    if not ip:
        return ''
    parts = str(ip).split('.')
    return f'{parts[0]}.{parts[1]}.*.*' if len(parts) == 4 else str(ip)


def _mask_url(url):
    if not url:
        return ''
    try:
        parsed = urlsplit(str(url).strip())
        hostname = (parsed.hostname or '').lower()
        allowed_host = (
            hostname == 'feishu.cn' or hostname.endswith('.feishu.cn') or
            hostname == 'larksuite.com' or hostname.endswith('.larksuite.com')
        )
        if parsed.scheme.lower() != 'https' or not parsed.netloc or not allowed_host:
            return ''
        clean = parsed._replace(query='', fragment='').geturl()
        return clean[:60] + '...' if len(clean) > 60 else clean
    except ValueError:
        return ''


def upsert_download_stat(base_dir, config, entry):
    values = _stat_values(entry)
    if not values['task_id']:
        seed = json.dumps(values, ensure_ascii=False, sort_keys=True).encode('utf-8')
        values['legacy_key'] = f"legacy_{values['ts']}_{hashlib.sha1(seed).hexdigest()[:16]}"
    else:
        values['legacy_key'] = None
    initialize_database(base_dir, config)
    connection = connect(base_dir, config)
    try:
        if values['task_id']:
            row = connection.execute(
                'SELECT * FROM download_stats WHERE task_id=?', (values['task_id'],)
            ).fetchone()
        else:
            row = connection.execute(
                'SELECT * FROM download_stats WHERE legacy_key=?', (values['legacy_key'],)
            ).fetchone()
        if row:
            old = dict(row)
            if values['ts'] < int(old.get('ts') or 0):
                return True
            merged = {
                field: (values[field] if values[field] else old.get(field) or '')
                for field in ('status', 'ts', 'time', 'url', 'path', 'title', 'ip', 'user')
            }
            connection.execute(
                """UPDATE download_stats SET status=?, ts=?, time=?, url=?, path=?,
                   title=?, ip=?, user=? WHERE row_id=?""",
                tuple(merged[field] for field in ('status', 'ts', 'time', 'url', 'path', 'title', 'ip', 'user'))
                + (old['row_id'],),
            )
        else:
            connection.execute(
                """INSERT INTO download_stats(
                   task_id, legacy_key, status, ts, time, url, path, title, ip, user
                ) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (values['task_id'], values['legacy_key'], values['status'], values['ts'], values['time'],
                 values['url'], values['path'], values['title'], values['ip'], values['user']),
            )
        connection.commit()
        return True
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def list_download_stats(base_dir, config, limit=None, page=None, page_size=None,
                        title='', url='', ip='', date=''):
    initialize_database(base_dir, config)
    connection = connect(base_dir, config)
    try:
        where = []
        params = []
        filters = (
            ('title', title, '(title LIKE ? COLLATE NOCASE OR user LIKE ? COLLATE NOCASE)'),
            ('url', url, 'url LIKE ? COLLATE NOCASE'),
            ('ip', ip, 'ip LIKE ? COLLATE NOCASE'),
        )
        for _, value, clause in filters:
            value = str(value or '').strip()
            if value:
                where.append(clause)
                like = f'%{value}%'
                params.extend((like, like) if _ == 'title' else (like,))
        date = str(date or '').strip()
        if date:
            try:
                start_date = datetime.strptime(date, '%Y-%m-%d')
                end_date = start_date + timedelta(days=1)
                where.append('time >= ? AND time < ?')
                params.extend((start_date.isoformat(timespec='seconds'), end_date.isoformat(timespec='seconds')))
            except ValueError:
                where.append('substr(time, 1, 10)=?')
                params.append(date)
        where_sql = (' WHERE ' + ' AND '.join(where)) if where else ''
        total = connection.execute(
            f'SELECT COUNT(*) AS count FROM download_stats{where_sql}', params
        ).fetchone()['count']
        sql = (
            'SELECT task_id AS id, status, ts, time, url, path, title, ip, user '
            f'FROM download_stats{where_sql} ORDER BY ts DESC, row_id DESC'
        )
        query_params = list(params)
        if page is not None or page_size is not None:
            try:
                current_page = max(int(page or 1), 1)
            except (TypeError, ValueError):
                current_page = 1
            try:
                current_page_size = min(max(int(page_size or 20), 1), 100)
            except (TypeError, ValueError):
                current_page_size = 20
            total_pages = (total + current_page_size - 1) // current_page_size if total else 0
            current_page = min(current_page, total_pages) if total_pages else 1
            sql += ' LIMIT ? OFFSET ?'
            query_params.extend((current_page_size, (current_page - 1) * current_page_size))
        elif limit:
            sql += ' LIMIT ?'
            query_params.append(min(max(int(limit), 1), 10000))
        rows = [dict(row) for row in connection.execute(sql, query_params).fetchall()]
        if page is not None or page_size is not None:
            return {
                'total': total,
                'items': rows,
                'page': current_page,
                'page_size': current_page_size,
                'total_pages': total_pages,
                'has_more': current_page * current_page_size < total,
            }
        return {'total': total, 'items': rows}
    finally:
        connection.close()


def delete_download_stats(base_dir, config, ts_list=None, id_list=None):
    ts_list = [str(value) for value in (ts_list or [])]
    id_list = [str(value) for value in (id_list or [])]
    if not ts_list and not id_list:
        return 0
    initialize_database(base_dir, config)
    connection = connect(base_dir, config)
    try:
        clauses = []
        params = []
        if id_list:
            placeholders = ','.join('?' for _ in id_list)
            clauses.append(f'task_id IN ({placeholders})')
            params.extend(id_list)
        if ts_list:
            placeholders = ','.join('?' for _ in ts_list)
            clauses.append(f"CAST(ts AS TEXT) IN ({placeholders})")
            params.extend(ts_list)
        cursor = connection.execute(
            f'DELETE FROM download_stats WHERE {" OR ".join(clauses)}', params
        )
        connection.commit()
        return cursor.rowcount
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def _legacy_paths(base_dir, config):
    log_dir = os.path.join(os.path.abspath(base_dir), str(config.get('log.dir', 'logs') or 'logs'))
    return os.path.join(log_dir, 'users.json'), os.path.join(log_dir, 'download_stats.jsonl')


def _backup_file(path):
    if not os.path.isfile(path):
        return None
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    target = f'{path}.migrated.{stamp}.bak'
    suffix = 1
    while os.path.exists(target):
        target = f'{path}.migrated.{stamp}.{suffix}.bak'
        suffix += 1
    shutil.copy2(path, target)
    _secure_file(target)
    return target


def _read_legacy_users(path):
    if not os.path.isfile(path):
        return []
    with open(path, 'r', encoding='utf-8') as handle:
        payload = json.load(handle)
    if not isinstance(payload, dict):
        return []
    records = []
    for open_id, record in payload.items():
        if isinstance(record, dict):
            record = dict(record)
            record.setdefault('open_id', open_id)
            records.append(record)
    return records


def _read_legacy_stats(path):
    if not os.path.isfile(path):
        return [], 0
    items = []
    invalid = 0
    with open(path, 'r', encoding='utf-8') as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            try:
                item = json.loads(line)
            except json.JSONDecodeError:
                invalid += 1
                continue
            if isinstance(item, dict):
                items.append(item)
            else:
                invalid += 1
    # 与旧 get_download_stats 保持一致：同一 task_id 保留最新记录并补齐旧记录字段。
    merged = {}
    for item in items:
        task_id = str(item.get('id') or '').strip()
        key = task_id or f"legacy_{item.get('ts')}"
        try:
            timestamp = int(item.get('ts') or 0)
        except (TypeError, ValueError):
            timestamp = 0
        old = merged.get(key)
        if old is None or timestamp >= int(old.get('ts') or 0):
            if old:
                for field, value in old.items():
                    if field not in item or not item[field]:
                        item[field] = value
            merged[key] = dict(item)
    return list(merged.values()), invalid


def migrate_legacy_data(base_dir, config, db_path=None, backup=True, dry_run=False, force=False):
    """迁移 users.json 和 download_stats.jsonl；默认幂等且不删除旧文件。"""
    users_path, stats_path = _legacy_paths(base_dir, config)
    users = []
    stats = []
    invalid_stats = 0
    users_marker = None
    stats_marker = None
    if not dry_run:
        initialize_database(base_dir, config, db_path=db_path)
        marker_connection = connect(base_dir, config, db_path=db_path)
        try:
            users_marker = _get_meta_from_connection(marker_connection, 'legacy_users_migrated')
            stats_marker = _get_meta_from_connection(marker_connection, 'legacy_stats_migrated')
        finally:
            marker_connection.close()
    if force or not users_marker:
        users = _read_legacy_users(users_path)
    if force or not stats_marker:
        stats, invalid_stats = _read_legacy_stats(stats_path)
    result = {
        'database': db_path or database_path(base_dir, config),
        'users_source': users_path,
        'stats_source': stats_path,
        'users_found': len(users),
        'stats_found': len(stats),
        'stats_invalid': invalid_stats,
        'users_imported': 0,
        'users_skipped': 0,
        'stats_imported': 0,
        'stats_skipped': 0,
        'backups': [],
        'dry_run': bool(dry_run),
    }
    if dry_run:
        return result

    connection = connect(base_dir, config, db_path=db_path)
    try:
        for record in users:
            values = _normalise_user(record)
            if not values:
                result['users_skipped'] += 1
                continue
            if users_marker and not force:
                result['users_skipped'] += 1
                continue
            existing = connection.execute('SELECT 1 FROM users WHERE open_id=?', (values['open_id'],)).fetchone()
            if existing and not force:
                result['users_skipped'] += 1
                continue
            connection.execute(
                """
                INSERT INTO users(
                    open_id, union_id, user_id, name, avatar, disabled, is_admin, is_system_admin,
                    access_token, refresh_token, token_expire_at,
                    refresh_token_expire_at, scope, token_invalid, created_at, last_login_at
                ) VALUES(
                    :open_id, :union_id, :user_id, :name, :avatar, :disabled, :is_admin, :is_system_admin,
                    :access_token, :refresh_token, :token_expire_at,
                    :refresh_token_expire_at, :scope, :token_invalid, :created_at, :last_login_at
                )
                ON CONFLICT(open_id) DO UPDATE SET
                    union_id=excluded.union_id, user_id=excluded.user_id, name=excluded.name,
                    avatar=excluded.avatar, disabled=excluded.disabled, is_admin=excluded.is_admin,
                    is_system_admin=excluded.is_system_admin,
                    access_token=excluded.access_token, refresh_token=excluded.refresh_token,
                    token_expire_at=excluded.token_expire_at,
                    refresh_token_expire_at=excluded.refresh_token_expire_at,
                    scope=excluded.scope, token_invalid=excluded.token_invalid,
                    created_at=excluded.created_at, last_login_at=excluded.last_login_at
                """,
                values,
            )
            result['users_imported'] += 1
        for item in stats:
            if stats_marker and not force:
                result['stats_skipped'] += 1
                continue
            # 使用同一连接导入，避免每条记录重复初始化数据库。
            values = _stat_values(item)
            if not values['task_id']:
                seed = json.dumps(values, ensure_ascii=False, sort_keys=True).encode('utf-8')
                values['legacy_key'] = f"legacy_{values['ts']}_{hashlib.sha1(seed).hexdigest()[:16]}"
            else:
                values['legacy_key'] = None
            existing = connection.execute(
                'SELECT * FROM download_stats WHERE task_id=?' if values['task_id'] else 'SELECT * FROM download_stats WHERE legacy_key=?',
                (values['task_id'],) if values['task_id'] else (values['legacy_key'],),
            ).fetchone()
            if existing and not force:
                result['stats_skipped'] += 1
                continue
            if existing:
                connection.execute(
                    """UPDATE download_stats SET status=?, ts=?, time=?, url=?, path=?,
                       title=?, ip=?, user=? WHERE row_id=?""",
                    (values['status'], values['ts'], values['time'], values['url'], values['path'],
                     values['title'], values['ip'], values['user'], existing['row_id']),
                )
            else:
                connection.execute(
                    """INSERT INTO download_stats(
                        task_id, legacy_key, status, ts, time, url, path, title, ip, user
                    ) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (values['task_id'], values['legacy_key'], values['status'], values['ts'], values['time'],
                     values['url'], values['path'], values['title'], values['ip'], values['user']),
                )
            result['stats_imported'] += 1
        if users and (not users_marker or force):
            _set_meta(connection, 'legacy_users_migrated', int(time.time()))
        if stats and (not stats_marker or force):
            _set_meta(connection, 'legacy_stats_migrated', int(time.time()))
        _set_meta(connection, 'schema_version', SCHEMA_VERSION)
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()

    if backup:
        for path in (users_path, stats_path):
            copied = _backup_file(path)
            if copied:
                result['backups'].append(copied)
    return result
