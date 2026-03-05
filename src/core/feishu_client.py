import os
import re
import threading
import time
import requests
from src.core.config_loader import ConfigLoader

class FeishuClient:

    def __init__(self, app_id, app_secret):
        self.app_id = app_id
        self.app_secret = app_secret
        self.logger = ConfigLoader.get_logger('feishu_client')
        self._token = ''
        self._expire_at = 0
        self._lock = threading.Lock()

    def get_token(self):
        now = int(time.time())
        with self._lock:
            if self._token and now < self._expire_at - 60:
                return self._token
            token = self._request_token()
            if token:
                return token
            return ''

    def _request_token(self):
        url = 'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal'
        payload = {'app_id': self.app_id, 'app_secret': self.app_secret}
        try:
            res = requests.post(url, json=payload, timeout=10).json()
        except Exception as e:
            self.logger.error('获取 Token 错误: ' + str(e))
            return ''
        if res.get('code') != 0:
            self.logger.error('获取 Token 失败: ' + str(res.get('msg', '')))
            return ''
        token = res.get('tenant_access_token') or ''
        expire = int(res.get('expire', 7200))
        self._token = token
        self._expire_at = int(time.time()) + expire
        return token

    def extract_doc_id(self, url):
        if not url:
            return ''
        m = re.search('/(docx|docs|wiki)/([a-zA-Z0-9]+)', url)
        if m and len(m.groups()) >= 2:
            return m.group(2)
        return ''

    def get_document_meta(self, doc_id):
        token = self.get_token()
        if not token:
            return {}
        url = f'https://open.feishu.cn/open-apis/docx/v1/documents/{doc_id}'
        headers = {'Authorization': 'Bearer ' + token}
        try:
            res = requests.get(url, headers=headers, timeout=10).json()
        except Exception as e:
            self.logger.error('获取文档元数据错误: ' + str(e))
            return {}
        if res.get('code') != 0:
            msg = res.get('msg', '')
            code = res.get('code')
            self.logger.error('获取文档元数据失败: ' + str(msg))
            if code == 99991668 or code == 99991663 or 'No permission' in msg or ('permission denied' in msg.lower()) or (code == 1770032):
                from src.core.config_loader import config
                bot_name = config.get('bot.name', 'Hawkon-Tool')
                contact_name = config.get('contact.name', 'Hakwon')
                error_msg = f'应用无权限，请为“{bot_name}”机器人开通管理权限，如有问题请联系 {contact_name}'
                raise PermissionError(error_msg)
            return {}
        return res.get('data', {}).get('document', {}) or {}

    def get_blocks(self, doc_id):
        token = self.get_token()
        if not token:
            return []
        url = f'https://open.feishu.cn/open-apis/docx/v1/documents/{doc_id}/blocks'
        headers = {'Authorization': 'Bearer ' + token}
        items = []
        page_token = ''
        while True:
            params = {'page_size': 500}
            if page_token:
                params['page_token'] = page_token
            try:
                res = requests.get(url, headers=headers, params=params, timeout=20).json()
            except Exception as e:
                self.logger.error(f'获取块请求错误: {str(e)}')
                raise RuntimeError(f'请求飞书接口失败: {str(e)}')
            code = res.get('code')
            if code != 0:
                msg = res.get('msg', '')
                self.logger.error(f'获取块 API 失败: code={code}, msg={msg}')
                if code == 1770032:
                    from src.core.config_loader import config
                    bot_name = config.get('bot.name', 'Hawkon-Tool')
                    contact_name = config.get('contact.name', 'Hakwon')
                    error_msg = f'应用无权限，请为 “{bot_name}” 机器人开通管理权限，如有问题请联系 ”{contact_name}“ 。'
                    raise PermissionError(error_msg)
                raise RuntimeError(f'飞书接口错误 ({code}): {msg}')
            data = res.get('data') or {}
            items.extend(data.get('items') or [])
            has_more = data.get('has_more')
            page_token = data.get('page_token') or ''
            if not has_more:
                break
        return items

    def download_media(self, file_token, save_path):
        token = self.get_token()
        if not token:
            return False
        url = f'https://open.feishu.cn/open-apis/drive/v1/medias/{file_token}/download'
        headers = {'Authorization': 'Bearer ' + token}
        try:
            r = requests.get(url, headers=headers, stream=True, timeout=30)
        except Exception as e:
            self.logger.error('下载媒体错误: ' + str(e))
            return False
        if r.status_code != 200:
            self.logger.error(f'下载媒体失败: {r.status_code}, 响应: {r.text[:200]}')
            if r.status_code == 403:
                from src.core.config_loader import config
                bot_name = config.get('bot.name', 'Hawkon-Tool')
                contact_name = config.get('contact.name', 'Hakwon')
                error_msg = f'下载图片/文件失败(403): 应用无权限，请为“{bot_name}”机器人开通【云文档】相关权限，如有问题请联系 {contact_name}'
                raise PermissionError(error_msg)
            if r.status_code == 400 and 'frequency limit' in r.text:
                self.logger.warning('触发频率限制，稍后重试...')
            return False
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        with open(save_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        return True

    def download_whiteboard(self, whiteboard_id, save_path):
        token = self.get_token()
        if not token:
            return False
        url = f'https://open.feishu.cn/open-apis/board/v1/whiteboards/{whiteboard_id}/download_as_image'
        headers = {'Authorization': 'Bearer ' + token}
        try:
            r = requests.get(url, headers=headers, stream=True, timeout=30)
        except Exception as e:
            self.logger.error('下载画板错误: ' + str(e))
            return False
        if r.status_code != 200:
            self.logger.error('下载画板失败: ' + str(r.status_code))
            return False
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        with open(save_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        return True

    def get_sheet_meta(self, spreadsheet_token, sheet_id=None):
        token = self.get_token()
        if not token:
            return {} if sheet_id else []
        url = f'https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/query'
        headers = {'Authorization': 'Bearer ' + token}
        try:
            res = requests.get(url, headers=headers, timeout=10).json()
            if res.get('code') == 0:
                sheets = res.get('data', {}).get('sheets', [])
                if sheet_id:
                    for sheet in sheets:
                        if sheet.get('sheet_id') == sheet_id:
                            return sheet
                    return {}
                return sheets
            self.logger.error(f"获取表格元数据失败: {res.get('msg')}")
        except Exception as e:
            self.logger.error('获取表格元数据错误: ' + str(e))
        return {} if sheet_id else []

    def get_sheet_values(self, spreadsheet_token, range_str):
        token = self.get_token()
        if not token:
            return {}
        url = f'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{range_str}'
        params = {'valueRenderOption': 'ToString', 'dateTimeRenderOption': 'FormattedString'}
        headers = {'Authorization': 'Bearer ' + token}
        try:
            res = requests.get(url, headers=headers, params=params, timeout=20).json()
            if res.get('code') == 0:
                return res.get('data', {}).get('valueRange', {})
            self.logger.error(f"获取表格数据失败: {res.get('msg')}")
        except Exception as e:
            self.logger.error('获取表格数据错误: ' + str(e))
        return {}

    def get_user_info(self, user_id):
        token = self.get_token()
        if not token:
            return {}
        url = f'https://open.feishu.cn/open-apis/contact/v3/users/{user_id}'
        headers = {'Authorization': 'Bearer ' + token}
        try:
            res = requests.get(url, headers=headers, timeout=10).json()
            if res.get('code') == 0:
                return res.get('data', {}).get('user', {})
            self.logger.error(f"获取用户信息失败: {res.get('msg')} ({res.get('code')})")
        except Exception as e:
            self.logger.error(f'获取用户信息错误: {str(e)}')
        return {}