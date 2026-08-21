import json
import os
import re
import shutil

def sanitize_name(name, ext=''):
    s = str(name or '')
    # 控制字符（\n \r \t 等）替换为空格，避免文件名/响应头含非法字符
    s = re.sub(r'[\x00-\x1f\x7f]', ' ', s)
    # 去除零宽字符
    s = re.sub('[\u200b\u200c\u200d\ufeff]', '', s)
    s = re.sub('[\\\\/:*?\\"<>|]', '_', s)
    # 压缩连续空白为单个空格
    s = re.sub(r'\s+', ' ', s).strip().strip('.')
    # 按 UTF-8 字节数截断至最多 150 字节（预留扩展名空间），保证不截断半个字符
    max_bytes = 150 - len(ext.encode('utf-8'))
    if len(s.encode('utf-8')) > max_bytes:
        s = s.encode('utf-8')[:max_bytes].decode('utf-8', errors='ignore').strip().strip('.')
    return (s or 'document') + ext

def safe_write(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'a', encoding='utf-8') as f:
        f.write(text)

def read_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def copy_file(src, dst):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copyfile(src, dst)