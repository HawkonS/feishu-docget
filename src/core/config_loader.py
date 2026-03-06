import os
import sys
import logging
from logging.handlers import RotatingFileHandler
CONFIG_META = [{'name': '飞书配置 (必填)', 'items': [{'key': 'feishu.app_id', 'default': '', 'desc': '# 飞书 App ID'}, {'key': 'feishu.app_secret', 'default': '', 'desc': '# 飞书 App Secret'}]}, {'name': '服务器配置', 'items': [{'key': 'server.port', 'default': '7800', 'desc': '# 服务器端口'}, {'key': 'admin.path', 'default': '/admin', 'desc': '# 管理后台路径'}, {'key': 'admin.password', 'default': 'Hawkon-FeishuDocGet', 'desc': '# 管理后台密码'}, {'key': 'template.default', 'default': 'Hawkon.docx', 'desc': '# 默认 Word 模板名称'}, {'key': 'template.upload_password', 'default': 'Hawkon-FeishuDocGet', 'desc': '# 模板上传密码'}]}, {'name': '页面显示配置', 'items': [{'key': 'page.title', 'default': '飞书文档下载工具', 'desc': '# 页面标题'}, {'key': 'page.description', 'default': '支持将飞书文档链接下载为指定模板的 Word 文件', 'desc': '# 页面描述'}, {'key': 'page.placeholder', 'default': '输入飞书文档链接，如 https://hawkon.feishu.cn/wiki/...', 'desc': '# 输入框占位符'}, {'key': 'page.usage_link_text', 'default': '使用说明', 'desc': '# 使用说明链接文本'}, {'key': 'url.usage_doc', 'default': 'https://github.com/HawkonS/feishu-docget', 'desc': '# 使用文档 URL'}, {'key': 'usage.url', 'default': 'mailto:contact@hawkon.tech', 'desc': '# 使用文档 URL / 联系方式'}, {'key': 'copyright.text', 'default': 'Hawkon 2025 -2026', 'desc': '# 版权文本'}, {'key': 'url.404', 'default': 'https://space.hawkon.tech/', 'desc': '# 404 重定向 URL'}, {'key': 'contact.name', 'default': 'Hakwon', 'desc': '# 联系人名称'}, {'key': 'bot.name', 'default': 'Hawkon-Tool', 'desc': '# 机器人名称'}]}, {'name': '导出设置', 'items': [{'key': 'image.max_height', 'default': '23', 'desc': '# 图片最大高度 (cm)'}, {'key': 'image.max_width', 'default': '16', 'desc': '# 图片最大宽度 (cm)'}, {'key': 'download.threads', 'default': '4', 'desc': '# 图片下载并发线程数'}, {'key': 'max.concurrent.downloads', 'default': '1', 'desc': '# 最大并发下载数'}, {'key': 'download_images', 'default': True, 'desc': '# 是否下载图片'}]}, {'name': '路径与日志配置', 'items': [{'key': 'workspace.dir', 'default': '.', 'desc': '# 工作区目录'}, {'key': 'template.dir', 'default': 'template', 'desc': '# 模板目录'}, {'key': 'output.dir', 'default': 'output', 'desc': '# 输出资源目录'}, {'key': 'output.max_size', 'default': '10G', 'desc': '# 输出目录最大大小'}, {'key': 'log.dir', 'default': 'logs', 'desc': '# 日志目录'}, {'key': 'log.level', 'default': 'INFO', 'desc': '# 日志级别'}, {'key': 'log.max_size', 'default': '20M', 'desc': '# 最大日志大小'}]}]
DEFAULT_CONFIG = {}
for group in CONFIG_META:
    for item in group['items']:
        DEFAULT_CONFIG[item['key']] = item['default']
CONFIG_FILE = 'feishu-docget.properties'

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

class ConfigLoader:
    _config = {}
    _initialized = False

    @classmethod
    def load_config(cls):
        if cls._initialized:
            return cls._config
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
        cls._write_config(config_path, cls._config)
        if not cls._config.get('feishu.app_id') or not cls._config.get('feishu.app_secret'):
            if sys.stdin.isatty():
                print('\n缺少飞书 App ID 或 App Secret。')
                if not cls._config.get('feishu.app_id'):
                    cls._config['feishu.app_id'] = input('请输入飞书 App ID: ').strip()
                if not cls._config.get('feishu.app_secret'):
                    cls._config['feishu.app_secret'] = input('请输入飞书 App Secret: ').strip()
                cls._write_config(config_path, cls._config)
            else:
                print('警告: 配置中缺少飞书 App ID 或 App Secret。')
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
        with open(path, 'w', encoding='utf-8') as f:
            for group in CONFIG_META:
                f.write('# ==========================================\n')
                f.write(f"# {group['name']}\n")
                f.write('# ==========================================\n')
                for item in group['items']:
                    key = item['key']
                    desc = item['desc']
                    if desc:
                        f.write(f'{desc}\n')
                    f.write(f"{key}={config.get(key, '')}\n")
                f.write('\n')

    @classmethod
    def get_comment(cls, key):
        comments = cls.get_comment_map()
        return comments.get(key, key)

    @classmethod
    def get_all_config_items(cls):
        cls.load_config()
        items = []
        for group in CONFIG_META:
            for item in group['items']:
                items.append({'key': item['key'], 'value': cls._config.get(item['key'], ''), 'desc': item['desc'], 'group': group['name']})
        return items

    @classmethod
    def save_config_from_admin(cls, new_config):
        config_path = os.path.join(os.getcwd(), CONFIG_FILE)
        cls._config.update(new_config)
        cls._write_config(config_path, cls._config)
        return True

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
        level_str = config.get('log.level', 'INFO').upper()
        level = getattr(logging, level_str, logging.INFO)
        logger = logging.getLogger(name)
        logger.setLevel(level)
        if not logger.handlers:
            handler = RotatingFileHandler(log_file, maxBytes=max_bytes, backupCount=5, encoding='utf-8')
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            console = logging.StreamHandler()
            console.setFormatter(formatter)
            logger.addHandler(console)
        return logger
config = ConfigLoader.load_config()