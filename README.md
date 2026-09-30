# 飞书文档下载器（feishu-docget）

<p align="center">
  <img src="./src/static/favicon.png" height="112" alt="feishu-docget 图标" />
</p>

<p align="center">
  把飞书云文档尽量保真地导出为 Word 文件
</p>

<p align="center">
  中文 ·
  <a href="./README_EN.md">English</a>
</p>

<p align="center">
  <a href="https://github.com/HawkonS/feishu-docget/actions/workflows/checks.yml"><img src="https://github.com/HawkonS/feishu-docget/actions/workflows/checks.yml/badge.svg" alt="Checks" /></a>
  <a href="https://github.com/HawkonS/feishu-docget/actions/workflows/codeql-analysis.yml"><img src="https://github.com/HawkonS/feishu-docget/actions/workflows/codeql-analysis.yml/badge.svg" alt="CodeQL" /></a>
  <a href="./LICENSE"><img src="https://img.shields.io/badge/license-Apache--2.0-blue.svg" alt="License: Apache-2.0" /></a>
</p>

## 简介

早期方案走的是 `feishu -> markdown -> docx` 中转路径，但复杂飞书文档转成 Markdown 时会丢失大量结构信息，例如合并单元格、富文本样式、图片/画板、嵌套列表等。本项目改为直接读取飞书开放平台返回的 Block 结构，用 `python-docx` 生成 Word 对象树，再套用模板并做格式清洗。

交付形态是一个 Web 服务（前台下载页 + 管理后台）和一个命令行工具。

## 功能

- 直出 Word：从飞书 Block 直接生成 `.docx`，减少 Markdown 中转造成的格式损失。
- 模板系统：上传和选择 `.docx` 模板，复用页眉、页脚、样式和封面。
- 格式可控：表格、图片/画板、代码块、正文段落、页边距、文档信息等都可按任务单独调整，规则不写回模板文件。
- 任务队列：前台提交后进入队列，可查看进度、日志和下载结果。
- 管理后台：项目、模板、配置、下载统计、日志与系统操作的统一入口。
- 命令行导出：`tools/feishu2word.sh`（Windows 为 `feishu2word.bat`）覆盖前台高级选项，便于脚本化。
- 自定义下载机器人：可临时指定飞书机器人凭据，校验通过后优先使用，权限不足时回退系统默认机器人。

## 快速开始

1. 在飞书开放平台创建自建应用，拿到 `App ID` 和 `App Secret`，开通云文档、电子表格、知识库、素材下载、画板下载等权限。企业内部文档还需要在文档页面把机器人加为协作者。
2. 执行 `./run.sh`（Windows 执行 `run.bat`）。首次启动会按 `requirements.txt` 安装依赖，并自动生成 `feishu-docget.properties`，在其中填入飞书凭据即可。
3. 打开 `http://127.0.0.1:7800/` 选模板、粘贴文档链接、创建任务；管理后台默认在 `http://127.0.0.1:7800/admin`。

导出结果写入 `output/<doc_id>/<文档标题>.docx`，图片放在同级 `img/` 目录，重复下载同一文档会尽量复用历史图片。

## 命令行

```bash
sh tools/feishu2word.sh "https://example.feishu.cn/wiki/xxxx" --template Hawkon.docx --style 3
```

参数与前台高级选项一一对应，完整列表执行 `sh tools/feishu2word.sh --help`；`--list-templates`、`--list-styles` 查看可用模板和表格样式，`--print-options` 只打印组装结果不下载。

## 配置

配置集中在项目根目录的 `feishu-docget.properties`，首次启动自动生成并补齐缺失项，同时创建日志、输出和模板目录。日常改配置建议直接在管理后台的配置管理页操作，页面上带有每项的说明和默认值；该文件包含敏感凭据，不要提交到 Git。

跨主机或容器部署、启用登录、反向代理等场景需要额外调整少数配置项，同样以配置管理页的说明为准。

## 常见问题

- **提示无权限**：检查应用权限范围、文档是否已把机器人加为协作者、链接是否属于当前企业空间。
- **图片或画板缺失**：确认开启了图片下载、应用具备素材下载权限，并查看日志中是否有媒体下载 403 或超时。
- **导出样式和模板不一致**：高级选项的清洗规则优先级高于模板，希望完全跟随模板时把相关选项留空或关闭。

## 开发

- 启用本地提交检查：`git config core.hooksPath .githooks`，之后每次提交都会跑单元测试（与 GitHub Actions 同一套）。
- 单元测试：`python -m unittest discover -s tests -p 'test_*.py'`；语法检查：`python -m compileall ./src`。
- 主要代码位置：Web/API 与后台在 `src/app.py`，转换与清洗在 `src/converters/docx/`，飞书接口在 `src/core/feishu_client.py`，页面在 `src/web/templates/`。

## 安全说明

- 不要提交 `feishu-docget.properties`、日志、导出文件和私有模板。
- 管理后台暴露了下载、删除、配置和系统操作能力，部署到公网前务必设置强密码，并放在可信网络或反向代理鉴权之后。
- 本项目仅供学习、归档和内部自动化场景使用，请遵守飞书平台规则和所在组织的数据合规要求。

## License

本项目依据 [Apache License 2.0](./LICENSE) 发布。
