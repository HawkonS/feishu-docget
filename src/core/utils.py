import json
import os
import re
import shutil

def sanitize_name(name):
    s = re.sub('[\\\\/:*?\\"<>|]', '_', str(name or ''))
    s = s.strip().strip('.')
    return s or 'document'

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