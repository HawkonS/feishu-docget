import os
import re
import stat
import sys
import time
import gzip
import shutil
import logging
import secrets
import tempfile
import threading
from logging.handlers import RotatingFileHandler

# ============================================================
# 配置 Schema（声明式定义，全部派生数据由其自动生成）
# 每项字段：
#   key              配置键名（与 properties 文件完全一致）
#   label            简短中文展示名（供前端 UI 使用）
#   desc             纯说明文本（写入配置文件时自动补回 '# ' 前缀）
#   default          默认值
#   type             str/int/float/bool/size/url/path/enum
#   options          enum 类型的可选值
#   min/max          数值型取值范围
#   unit             单位（如 cm）
#   sensitive        是否敏感（API 返回时脱敏）
#   restart_required 修改后是否需要重启服务才能生效
# 可选元字段（仅按需声明，前端与校验自动消费）：
#   gen              键级"生成"辅助按钮类型（如 secret-hex-64）
#   placeholder      输入框占位提示
#   hint             格式补充说明
#   pattern          校验正则（validate_config 与前端共用单源）
#   pattern_ignorecase  pattern 是否忽略大小写
#   depends_on       依赖的开关键（仅当该键启用时本项才生效）
#
# 契约：新增配置项（type 为已有类型）只需在 CONFIG_SCHEMA 加一项，
# 前端控件/校验/提交、API 白名单、文件写盘全部自动生效，无需改动其他代码。
# 需要前端感知的仅两种情形：新增全新 type、通过 gen 字段声明的键级辅助按钮。
# ============================================================
CONFIG_SCHEMA = [
    {
        'name': '飞书配置 (必填)',
        'items': [
            {
                'key': 'feishu.app_id', 'label': '飞书 App ID', 'type': 'str',
                'default': '', 'desc': '飞书 App ID',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'feishu.app_secret', 'label': '飞书 App Secret', 'type': 'str',
                'default': '', 'desc': '飞书 App Secret',
                'sensitive': True, 'restart_required': False,
            },
            {
                'key': 'login.enabled', 'label': '启用登录', 'type': 'bool',
                'default': 'false',
                'desc': '是否启用登录机制（飞书 OAuth），开启后未登录不能访问主页 (true/false)，开启后下方 OAuth 回调地址生效',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'login.oauth.redirect_uri', 'label': 'OAuth 回调地址', 'type': 'str',
                'default': '',
                'desc': '飞书 OAuth 回调地址（选填，留空时按访问域名自动生成，启用 server.https.enabled 时自动使用 https）。仅在启用登录时生效；手动填写必须与飞书开放平台『安全设置-重定向 URL』及授权时使用的地址完全一致',
                'depends_on': 'login.enabled',
                'sensitive': False, 'restart_required': False,
            },
        ],
    },
    {
        'name': '服务器配置',
        'items': [
            {
                'key': 'server.port', 'label': '服务器端口', 'type': 'int',
                'default': '7800', 'min': 1, 'max': 65535,
                'pattern': r'^[+-]?\d+$',
                'desc': '服务器端口',
                'sensitive': False, 'restart_required': True,
            },
            {
                'key': 'server.https.enabled', 'label': '启用 HTTPS', 'type': 'bool',
                'default': 'false',
                'desc': '是否启用 HTTPS，适用于反向代理 SSL 终结场景 (true/false)',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'admin.path', 'label': '管理后台路径', 'type': 'str',
                'default': '/admin', 'desc': '管理后台路径',
                'sensitive': False, 'restart_required': True,
            },
            {
                'key': 'admin.password', 'label': '管理后台密码', 'type': 'str',
                'default': '', 'desc': '管理后台密码 (留空则首次启动时自动生成随机密码)',
                'sensitive': True, 'restart_required': False,
            },
            {
                'key': 'system.sudo_password', 'label': '系统 sudo 密码', 'type': 'str',
                'default': '', 'desc': '系统 sudo 密码 (用于自动更新脚本，选填)',
                'sensitive': True, 'restart_required': False,
            },
            {
                'key': 'server.secret_key', 'label': 'Session 密钥', 'type': 'str',
                'default': '', 'desc': '用于 Flask session 加密的密钥，留空则自动生成',
                'gen': 'secret-hex-64',
                'sensitive': True, 'restart_required': True,
            },
            {
                'key': 'server.threads', 'label': '服务器线程数', 'type': 'int',
                'default': '8', 'min': 1,
                'pattern': r'^[+-]?\d+$',
                'desc': 'waitress 生产服务器工作线程数',
                'sensitive': False, 'restart_required': True,
            },
        ],
    },
    {
        'name': '页面显示配置',
        'items': [
            {
                'key': 'page.title', 'label': '页面标题', 'type': 'str',
                'default': '飞书文档下载工具', 'desc': '页面标题',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'page.favicon', 'label': '网站图标路径', 'type': 'path',
                'default': 'src/static/favicon.ico',
                'desc': '网站图标文件路径 (相对工作区目录，支持 ico/png/svg)',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'page.description', 'label': '页面描述', 'type': 'str',
                'default': '支持将飞书文档链接下载为指定模板的 Word 文件', 'desc': '页面描述',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'page.placeholder', 'label': '输入框占位符', 'type': 'str',
                'default': '输入飞书文档链接，如 https://hawkon.feishu.cn/wiki/...',
                'desc': '输入框占位符',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'page.usage_link_text', 'label': '使用说明链接文本', 'type': 'str',
                'default': '使用说明', 'desc': '使用说明链接文本',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'url.usage_doc', 'label': '使用文档链接', 'type': 'url',
                'default': 'https://github.com/HawkonS/feishu-docget',
                'placeholder': '以 http(s):// 或 mailto: 开头', 'pattern': r'^(https?://|mailto:)',
                'desc': '使用文档链接',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'usage.url', 'label': '联系方式', 'type': 'url',
                'default': 'mailto:contact@hawkon.tech',
                'placeholder': '以 http(s):// 或 mailto: 开头', 'pattern': r'^(https?://|mailto:)',
                'desc': '联系方式（支持 mailto: 或 http 链接）',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'copyright.text', 'label': '版权文本', 'type': 'str',
                'default': 'Hawkon 2025 -2026', 'desc': '版权文本',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'url.404', 'label': '404 重定向 URL', 'type': 'url',
                'default': 'https://space.hawkon.tech/', 'desc': '404 重定向 URL',
                'placeholder': '以 http(s):// 或 mailto: 开头', 'pattern': r'^(https?://|mailto:)',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'contact.name', 'label': '联系人名称', 'type': 'str',
                'default': 'Hakwon', 'desc': '联系人名称',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'bot.name', 'label': '机器人名称', 'type': 'str',
                'default': 'Hawkon-Tool', 'desc': '机器人名称',
                'sensitive': False, 'restart_required': False,
            },
        ],
    },
    {
        'name': '导出设置',
        'items': [
            {
                'key': 'template.default', 'label': '默认 Word 模板', 'type': 'str',
                'default': 'Hawkon.docx', 'desc': '默认 Word 模板名称',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'image.max_height', 'label': '图片最大高度', 'type': 'float',
                'default': '23', 'min': 0.1, 'unit': 'cm',
                'desc': '图片最大高度 (cm)',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'image.max_width', 'label': '图片最大宽度', 'type': 'float',
                'default': '16', 'min': 0.1, 'unit': 'cm',
                'desc': '图片最大宽度 (cm)',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'download.threads', 'label': '图片下载线程数', 'type': 'int',
                'default': '4', 'min': 1,
                'pattern': r'^[+-]?\d+$',
                'desc': '图片下载并发线程数',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'max.concurrent.downloads', 'label': '最大并发下载数', 'type': 'int',
                'default': '1', 'min': 1,
                'pattern': r'^[+-]?\d+$',
                'desc': '最大并发下载数；个人登录任务按飞书账号分别限制，机器人/未登录任务共享此上限',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'download_images', 'label': '下载图片', 'type': 'bool',
                'default': 'true', 'desc': '是否下载图片',
                'sensitive': False, 'restart_required': False,
            },
        ],
    },
    {
        'name': '路径与日志配置',
        'items': [
            {
                'key': 'workspace.dir', 'label': '工作区目录', 'type': 'path',
                'default': '.', 'desc': '工作区目录',
                'sensitive': False, 'restart_required': True,
            },
            {
                'key': 'template.dir', 'label': '模板目录', 'type': 'path',
                'default': 'template', 'desc': '模板目录',
                'sensitive': False, 'restart_required': True,
            },
            {
                'key': 'template.password.long_term', 'label': '模板上传密码（长期存储）', 'type': 'str',
                'default': '', 'desc': '模板上传密码 - 长期存储模式',
                'sensitive': True, 'restart_required': False,
            },
            {
                'key': 'template.password.one_time', 'label': '模板上传密码（仅本次使用）', 'type': 'str',
                'default': '', 'desc': '模板上传密码 - 仅本次使用模式',
                'sensitive': True, 'restart_required': False,
            },
            {
                'key': 'output.dir', 'label': '输出资源目录', 'type': 'path',
                'default': 'output', 'desc': '输出资源目录',
                'sensitive': False, 'restart_required': True,
            },
            {
                'key': 'output.max_size', 'label': '输出目录最大大小', 'type': 'size',
                'default': '10G', 'desc': '输出目录最大大小',
                'placeholder': '如 10G / 20M / 512K', 'pattern': r'^\d+[KMG]?$',
                'pattern_ignorecase': True, 'hint': '数字+可选单位 K/M/G',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'log.dir', 'label': '日志目录', 'type': 'path',
                'default': 'logs', 'desc': '日志目录',
                'sensitive': False, 'restart_required': True,
            },
            {
                'key': 'log.level', 'label': '日志级别', 'type': 'enum',
                'default': 'INFO',
                'options': ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
                'desc': '日志级别',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'log.max_size', 'label': '最大日志大小', 'type': 'size',
                'default': '20M', 'desc': '最大日志大小',
                'placeholder': '如 10G / 20M / 512K', 'pattern': r'^\d+[KMG]?$',
                'pattern_ignorecase': True, 'hint': '数字+可选单位 K/M/G',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'log.backup_count', 'label': '日志轮转保留数量', 'type': 'int',
                'default': '5', 'min': 0,
                'pattern': r'^[+-]?\d+$',
                'desc': '日志轮转保留的历史文件数量',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'log.compress', 'label': '历史日志压缩', 'type': 'bool',
                'default': 'true',
                'desc': '轮转出的历史日志是否压缩为 .gz (true/false)',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'log.archive.enabled', 'label': '历史日志归档', 'type': 'bool',
                'default': 'true',
                'desc': '超出保留数量的历史日志是否归档保存，false 则直接删除 (true/false)',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'log.archive.dir', 'label': '日志归档目录', 'type': 'path',
                'default': '', 'desc': '日志归档目录 (留空则使用 <日志目录>/archive)',
                'sensitive': False, 'restart_required': False,
            },
            {
                'key': 'log.archive.max_days', 'label': '归档日志保留天数', 'type': 'int',
                'default': '30', 'min': 0,
                'pattern': r'^[+-]?\d+$',
                'desc': '归档日志保留天数 (0 表示永久保留)',
                'sensitive': False, 'restart_required': False,
            },
        ],
    },
]

# key -> Schema item 索引
SCHEMA_ITEM_MAP = {item['key']: item for group in CONFIG_SCHEMA for item in group['items']}

# 兼容旧结构：CONFIG_META（desc 带 '# ' 前缀），供注释写盘等逻辑使用
CONFIG_META = [
    {'name': group['name'],
     'items': [{'key': item['key'], 'default': item['default'], 'desc': '# ' + item['desc']}
               for item in group['items']]}
    for group in CONFIG_SCHEMA
]

# 文件写入顺序：与存量 properties 文件的区块布局保持一致。
# 注意 template.default 在 UI 分组上属于"导出设置"，但文件中仍写在"服务器配置"区块；
# login.enabled 在 UI 分组上属于"飞书配置 (必填)"，但文件中仍写在原"服务器配置"区块位置。
# 调整分组只影响注释展示与前端分组，不影响键的读写。
FILE_WRITE_ORDER = [
    {'name': group['name'],
     'items': [SCHEMA_ITEM_MAP[item['key']] for item in group['items']]}
    for group in CONFIG_SCHEMA
]
for _group in FILE_WRITE_ORDER:
    if _group['name'] == '服务器配置':
        _group['items'].insert(4, SCHEMA_ITEM_MAP['template.default'])
    elif _group['name'] == '导出设置':
        _group['items'] = [i for i in _group['items'] if i['key'] != 'template.default']
    elif _group['name'] == '飞书配置 (必填)':
        # login.enabled 仅 UI 分组归拢，写盘位置保持存量文件原布局（服务器配置区块）
        _group['items'] = [i for i in _group['items'] if i['key'] != 'login.enabled']
for _group in FILE_WRITE_ORDER:
    if _group['name'] == '服务器配置':
        # 插回 server.threads 之前，保持存量文件中 secret_key → login.enabled 的原相对顺序
        _idx = next(i for i, _it in enumerate(_group['items']) if _it['key'] == 'server.threads')
        _group['items'].insert(_idx, SCHEMA_ITEM_MAP['login.enabled'])

DEFAULT_CONFIG = {item['key']: item['default'] for item in SCHEMA_ITEM_MAP.values()}
# 配置键白名单：仅允许写入 Schema 中定义的键，防止配置注入
ALLOWED_CONFIG_KEYS = set(DEFAULT_CONFIG.keys())
# 敏感配置键：由 Schema 的 sensitive 字段自动派生（对外导出名称保持兼容）
SENSITIVE_KEYS = {key for key, item in SCHEMA_ITEM_MAP.items() if item.get('sensitive')}

CONFIG_FILE = 'feishu-docget.properties'

# 保护配置文件写路径（_write_config 与 save_config_from_admin 共用，可重入以支持嵌套加锁）
_config_write_lock = threading.RLock()

_lazy_logger = None

def _get_logger():
    """延迟创建独立 logger，避免在配置加载早期产生依赖。"""
    global _lazy_logger
    if _lazy_logger is None:
        _lazy_logger = logging.getLogger('config_loader')
    return _lazy_logger

def parse_size(size_str):
    size_str = size_str.upper()
    if size_str.endswith('K'):
        return int(size_str[:-1]) * 1024
    elif size_str.endswith('M'):
        return int(size_str[:-1]) * 1024 * 1024
    elif size_str.endswith('G'):
        return int(size_str[:-1]) * 1024 * 1024 * 1024
    else:
        return int(size_str)

def parse_bool(value, default=False):
    if value is None or str(value).strip() == '':
        return default
    return str(value).strip().lower() in ('true', '1', 'yes', 'on')

_SIZE_PATTERN = re.compile(r'^\d+[KMG]?$', re.IGNORECASE)

def validate_config(updates):
    """按 Schema 逐项校验配置变更（仅由 API 层调用）。

    - 空字符串对任何项放行（保留"留空"语义，如 admin.password/log.archive.dir）
    - Schema 项声明了 pattern 时，格式校验以 pattern 为单源（与前端共用）；
      未声明 pattern 的项走原 type 分支语义
    - 返回错误列表 [{'key': ..., 'message': ...}]，全部合法时返回空列表
    """
    errors = []
    for key, value in (updates or {}).items():
        item = SCHEMA_ITEM_MAP.get(key)
        if item is None:
            errors.append({'key': key, 'message': '未知的配置项'})
            continue
        if value is None or str(value).strip() == '':
            continue
        t = item.get('type', 'str')
        s = str(value).strip()
        # pattern 单源：url 为前缀匹配，size/int 的 pattern 已含 ^$ 为完整匹配，均用 re.match
        pattern = item.get('pattern')
        if pattern:
            flags = re.IGNORECASE if item.get('pattern_ignorecase') else 0
            if not re.match(pattern, s, flags):
                errors.append({'key': key, 'message': f'格式不正确，当前值: {s}'})
                continue
        if t == 'int':
            try:
                num = int(s)
            except (TypeError, ValueError):
                errors.append({'key': key, 'message': f'需要整数，当前值: {s}'})
                continue
            if not _in_range(num, item, key, errors):
                continue
        elif t == 'float':
            try:
                num = float(s)
            except (TypeError, ValueError):
                errors.append({'key': key, 'message': f'需要数字，当前值: {s}'})
                continue
            if not _in_range(num, item, key, errors):
                continue
        elif t == 'bool':
            if isinstance(value, bool):
                continue
            if s.lower() not in ('true', 'false'):
                errors.append({'key': key, 'message': f'仅接受 true 或 false，当前值: {s}'})
        elif t == 'size' and not pattern:
            if not _SIZE_PATTERN.match(s):
                errors.append({'key': key, 'message': f'大小格式应为 数字+可选单位 K/M/G，当前值: {s}'})
        elif t == 'url' and not pattern:
            if not (s.startswith('http://') or s.startswith('https://') or s.startswith('mailto:')):
                errors.append({'key': key, 'message': f'需以 http(s):// 或 mailto: 开头，当前值: {s}'})
        elif t == 'enum':
            options = item.get('options') or []
            if not any(s == str(o) or s.upper() == str(o).upper() for o in options):
                errors.append({'key': key, 'message': f'可选值: {"/".join(str(o) for o in options)}，当前值: {s}'})
        # str/path 类型不做格式校验
    return errors

def _in_range(num, item, key, errors):
    """校验数值是否在 Schema 的 min/max 范围内，返回是否通过。"""
    lo = item.get('min')
    hi = item.get('max')
    if lo is not None and num < lo:
        msg = f'不能小于 {lo}'
        if hi is not None:
            msg = f'需在 {lo}-{hi} 范围内'
        errors.append({'key': key, 'message': f'{msg}，当前值: {num}'})
        return False
    if hi is not None and num > hi:
        errors.append({'key': key, 'message': f'不能大于 {hi}，当前值: {num}'})
        return False
    return True

class ArchivingRotatingFileHandler(RotatingFileHandler):
    """支持压缩与归档的轮转日志 Handler。

    - compress: 轮转出的历史日志压缩为 .gz
    - archive_dir: 超出 backupCount 的最旧历史日志移动到归档目录，而不是直接删除
    - archive_max_days: 归档文件按修改时间清理，超过天数的自动删除 (0 表示不清理)
    """

    def __init__(self, filename, maxBytes=0, backupCount=0, encoding=None,
                 compress=True, archive_dir=None, archive_max_days=0):
        self.compress = compress
        self.archive_dir = archive_dir
        self.archive_max_days = archive_max_days
        super().__init__(filename, maxBytes=maxBytes, backupCount=backupCount, encoding=encoding)

    def _log_stem(self):
        base = os.path.basename(self.baseFilename)
        return base[:-4] if base.endswith('.log') else base

    def rotation_filename(self, default_name):
        if not self.compress:
            return default_name
        return default_name + '.gz'

    def rotate(self, source, dest):
        if not self.compress:
            super().rotate(source, dest)
            return
        if os.path.exists(source):
            with open(source, 'rb') as sf, gzip.open(dest, 'wb') as df:
                shutil.copyfileobj(sf, df)
            os.remove(source)

    def doRollover(self):
        if self.archive_dir:
            oldest = self.rotation_filename("%s.%d" % (self.baseFilename, self.backupCount))
            if os.path.exists(oldest):
                self._archive_file(oldest)
        super().doRollover()
        if self.archive_dir:
            self._cleanup_old_archives()

    def _archive_file(self, path):
        try:
            os.makedirs(self.archive_dir, exist_ok=True)
            stem = self._log_stem()
            suffix = '.log.gz' if path.endswith('.gz') else '.log'
            dest = os.path.join(self.archive_dir, f'{stem}-{time.strftime("%Y%m%d-%H%M%S")}{suffix}')
            seq = 1
            while os.path.exists(dest):
                dest = os.path.join(self.archive_dir, f'{stem}-{time.strftime("%Y%m%d-%H%M%S")}-{seq}{suffix}')
                seq += 1
            shutil.move(path, dest)
        except Exception as e:
            self.handleError(logging.LogRecord('', 0, '', 0, f'日志归档失败 {path}: {e}', None, None))

    def _cleanup_old_archives(self):
        if not self.archive_max_days:
            return
        try:
            cutoff = time.time() - self.archive_max_days * 86400
            prefix = self._log_stem() + '-'
            for name in os.listdir(self.archive_dir):
                if not name.startswith(prefix):
                    continue
                fp = os.path.join(self.archive_dir, name)
                if os.path.isfile(fp) and os.path.getmtime(fp) < cutoff:
                    os.remove(fp)
        except Exception:
            pass

class ConfigLoader:
    _config = {}
    _initialized = False

    @classmethod
    def load_config(cls):
        if cls._initialized:
            return cls._config
        changed = False
        config_path = os.path.join(os.getcwd(), CONFIG_FILE)
        if os.path.exists(config_path):
            cls._config = cls._read_config(config_path)
        else:
            print(f'配置文件 {CONFIG_FILE} 未找到。正在创建默认配置...')
            cls._write_config(config_path, DEFAULT_CONFIG)
            cls._config = DEFAULT_CONFIG.copy()
        for k, v in DEFAULT_CONFIG.items():
            if k not in cls._config:
                cls._config[k] = v
                changed = True
        # 管理后台密码为空时自动生成高熵随机密码，避免使用可预测的固定密码
        if not str(cls._config.get('admin.password') or '').strip():
            generated_password = secrets.token_urlsafe(16)
            cls._config['admin.password'] = generated_password
            changed = True
            print(f'已自动生成管理后台密码并写入 {CONFIG_FILE}: {generated_password}')
            print('请及时登录管理后台修改为自定义密码，并妥善保管。')
        if not cls._config.get('feishu.app_id') or not cls._config.get('feishu.app_secret'):
            if sys.stdin.isatty():
                print('\n缺少飞书 App ID 或 App Secret。')
                if not cls._config.get('feishu.app_id'):
                    cls._config['feishu.app_id'] = input('请输入飞书 App ID: ').strip()
                if not cls._config.get('feishu.app_secret'):
                    cls._config['feishu.app_secret'] = input('请输入飞书 App Secret: ').strip()
                changed = True
            else:
                print('警告: 配置中缺少飞书 App ID 或 App Secret。')
        # 仅在确有变更时写盘，避免无条件全量重写
        if changed:
            cls._write_config(config_path, cls._config)
        workspace = cls._config.get('workspace.dir', '.')
        log_dir = os.path.join(workspace, cls._config.get('log.dir', 'logs'))
        output_dir = os.path.join(workspace, cls._config.get('output.dir', 'output'))
        template_dir = os.path.join(workspace, cls._config.get('template.dir', 'template'))
        os.makedirs(log_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)
        os.makedirs(template_dir, exist_ok=True)
        cls._initialized = True
        return cls._config

    @classmethod
    def _read_config(cls, path):
        config = {}
        try:
            with open(path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                    if '=' in line:
                        k, v = line.split('=', 1)
                        config[k.strip()] = v.strip()
        except Exception as e:
            print(f'读取配置错误 {path}: {e}')
        return config

    @classmethod
    def get_comment_map(cls):
        comments = {}
        for group in CONFIG_META:
            for item in group['items']:
                comments[item['key']] = item['desc']
        return comments

    @classmethod
    def _write_config(cls, path, config):
        """原子写配置文件：先写同目录临时文件，成功后 os.replace 替换。

        成功返回 True，失败返回 False；若目标为 symlink 则透过链接写真实文件，
        并在替换后尽力恢复原文件的权限与属主。
        """
        with _config_write_lock:
            # 若目标是 symlink，解析到真实文件，避免 replace 把链接替换成普通文件
            real_path = os.path.realpath(path)
            directory = os.path.dirname(real_path) or '.'
            # 记录原文件 stat，替换后用于恢复权限/属主
            old_stat = None
            try:
                old_stat = os.stat(real_path)
            except OSError:
                pass
            tmp_path = None
            try:
                fd, tmp_path = tempfile.mkstemp(prefix='.feishu-docget.properties.', suffix='.tmp', dir=directory)
                with os.fdopen(fd, 'w', encoding='utf-8') as f:
                    for group in FILE_WRITE_ORDER:
                        f.write('# ==========================================\n')
                        f.write(f"# {group['name']}\n")
                        f.write('# ==========================================\n')
                        for item in group['items']:
                            key = item['key']
                            desc = item.get('desc', '')
                            if desc:
                                f.write(f'# {desc}\n')
                            val = config.get(key, item['default'])
                            if val is None: val = ''
                            f.write(f"{key}={val}\n")
                        f.write('\n')
                    f.flush()
                    os.fsync(f.fileno())
                os.replace(tmp_path, real_path)
                tmp_path = None
                # mkstemp 固定 0600 且 replace 换 inode，需恢复原文件权限与属主
                if old_stat is not None:
                    os.chmod(real_path, stat.S_IMODE(old_stat.st_mode))
                    try:
                        os.chown(real_path, old_stat.st_uid, old_stat.st_gid)
                    except (PermissionError, OSError):
                        pass
                return True
            except PermissionError as e:
                print(f"警告: 无法写入配置文件 {path}: {e}")
                print("请检查文件是否被设置为只读、隐藏，或正在被其他程序使用。")
            except Exception as e:
                print(f"写入配置文件 {path} 失败: {e}")
            finally:
                if tmp_path and os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except Exception:
                        pass
            return False

    @classmethod
    def get_comment(cls, key):
        comments = cls.get_comment_map()
        return comments.get(key, key)

    @classmethod
    def get_meta_map(cls):
        """返回 key -> Schema item 的字典。"""
        return dict(SCHEMA_ITEM_MAP)

    @classmethod
    def get_int(cls, key, default=0):
        try:
            return int(str(cls.load_config().get(key, default)).strip())
        except (TypeError, ValueError):
            _get_logger().warning(f'配置项 {key} 无法解析为整数，回退默认值: {default}')
            return default

    @classmethod
    def get_float(cls, key, default=0.0):
        try:
            return float(str(cls.load_config().get(key, default)).strip())
        except (TypeError, ValueError):
            _get_logger().warning(f'配置项 {key} 无法解析为数字，回退默认值: {default}')
            return default

    @classmethod
    def get_bool(cls, key, default=False):
        return parse_bool(cls.load_config().get(key), default)

    @classmethod
    def get_size(cls, key, default=0):
        try:
            return parse_size(str(cls.load_config().get(key, default)).strip())
        except (TypeError, ValueError):
            _get_logger().warning(f'配置项 {key} 无法解析为大小，回退默认值: {default}')
            return default

    @classmethod
    def get_all_config_items(cls):
        cls.load_config()
        items = []
        for group in CONFIG_SCHEMA:
            for meta in group['items']:
                key = meta['key']
                raw_value = cls._config.get(key, '')
                sensitive = bool(meta.get('sensitive'))
                masked = sensitive and bool(raw_value)
                items.append({
                    'key': key,
                    'value': raw_value,
                    'desc': meta['desc'],
                    'group': group['name'],
                    'label': meta.get('label', key),
                    'type': meta.get('type', 'str'),
                    'options': meta.get('options'),
                    'min': meta.get('min'),
                    'max': meta.get('max'),
                    'unit': meta.get('unit'),
                    'sensitive': sensitive,
                    'masked': masked,
                    'restart_required': bool(meta.get('restart_required')),
                    'gen': meta.get('gen'),
                    'placeholder': meta.get('placeholder'),
                    'hint': meta.get('hint'),
                    'pattern': meta.get('pattern'),
                    'pattern_ignorecase': meta.get('pattern_ignorecase'),
                    'depends_on': meta.get('depends_on'),
                })
        for item in items:
            if item['masked']:
                item['value'] = '******'
        return items

    @classmethod
    def save_config_from_admin(cls, new_config):
        config_path = os.path.join(os.getcwd(), CONFIG_FILE)
        # 仅接受白名单内的配置键，拒绝任意键注入
        # 注意：本方法不做校验（值可能由程序生成），校验由 API 层 validate_config 负责
        filtered = {k: v for k, v in new_config.items() if k in ALLOWED_CONFIG_KEYS}
        def _norm(v):
            return '' if v is None else str(v).strip()
        with _config_write_lock:
            # 无实质变更时短路：filtered 为空或所有值与当前配置逐一相等，
            # 跳过 update 与写盘，避免无意义地刷新文件 mtime 与注释
            if not filtered or all(_norm(v) == _norm(cls._config.get(k)) for k, v in filtered.items()):
                return True
            cls._config.update(filtered)
            return cls._write_config(config_path, cls._config)

    @classmethod
    def get_logger(cls, name, filename=None):
        config = cls.load_config()
        workspace = config.get('workspace.dir', '.')
        log_dir = os.path.join(workspace, config.get('log.dir', 'logs'))
        os.makedirs(log_dir, exist_ok=True)
        if not filename:
            filename = f'{name}.log'
        log_file = os.path.join(log_dir, filename)
        max_bytes = parse_size(config.get('log.max_size', '20M'))
        try:
            backup_count = int(config.get('log.backup_count', '5'))
        except (TypeError, ValueError):
            backup_count = 5
        compress = parse_bool(config.get('log.compress', 'true'), True)
        archive_enabled = parse_bool(config.get('log.archive.enabled', 'true'), True)
        archive_dir = config.get('log.archive.dir', '').strip() or os.path.join(log_dir, 'archive')
        try:
            archive_max_days = int(config.get('log.archive.max_days', '30'))
        except (TypeError, ValueError):
            archive_max_days = 30
        level_str = config.get('log.level', 'INFO').upper()
        level = getattr(logging, level_str, logging.INFO)
        logger = logging.getLogger(name)
        logger.setLevel(level)
        if not logger.handlers:
            handler = ArchivingRotatingFileHandler(
                log_file, maxBytes=max_bytes, backupCount=backup_count, encoding='utf-8',
                compress=compress,
                archive_dir=archive_dir if archive_enabled else None,
                archive_max_days=archive_max_days)
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            console = logging.StreamHandler()
            console.setFormatter(formatter)
            logger.addHandler(console)
        return logger
config = ConfigLoader.load_config()
