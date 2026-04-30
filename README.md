# feishu-docget

飞书文档下载工具。项目面向“把飞书云文档尽量保真地导出为 Word 文件”的场景，提供 Web 页面、管理后台和命令行入口，支持 Word 模板、表格样式、图片处理、代码块样式、文档信息写入、下载统计和日志查看。

## 背景

早期方案尝试过 `feishu -> markdown -> docx` 的中转路径，但复杂飞书文档在转成 Markdown 时会丢失大量结构信息，例如合并单元格、富文本样式、图片/画板、嵌套列表等。本项目改为直接读取飞书开放平台返回的 Block 结构，并用 `python-docx` 生成 Word 对象树，再做模板和格式清洗。

当前重点能力是 Word 导出。Markdown/HTML 等格式暂未作为主线维护。

## 主要功能

- 直出 Word：直接从飞书文档 Block 生成 `.docx`，减少 Markdown 中转造成的格式损失。
- 模板系统：支持上传和选择 `.docx` 模板，可复用页眉、页脚、样式和封面。
- 表格样式：前台提供 6 种表格样式，导出时按 Word OOXML 写入边框、底色和文字色。
- 自定义表格边框：可在高级选项中启用统一颜色，并分别设置上下左右边框的线型和粗细。
- 图片和画板：支持图片下载、最大宽高限制、对齐方式设置，画板会下载为图片并裁剪空白。
- 代码块样式：代码块以 Word 表格承载，可设置背景色、字体、字号、对齐、边框和宽度。
- 高级格式清洗：支持正文段落样式、图片样式、表格布局、页边距、标题编号清理等。
- 文档信息写入：可写入作者、标题、主题、创建/修改时间、公司、模板等 Word 元数据。
- 任务管理：Web 前台支持任务队列、进度、日志和下载。
- 管理后台：支持项目管理、模板维护、配置管理、下载统计、日志检查和系统操作。
- 命令行导出：`src/cli/feishu2word.py` 可用于脚本化下载。

## 技术栈

- Python 3
- Flask
- requests
- python-docx
- lxml
- Pillow

项目没有单独的 `requirements.txt`，`run.bat` 和 `run.sh` 会尝试安装上述依赖。手动安装可执行：

```bash
pip install Flask requests python-docx lxml Pillow
```

## 项目结构

```text
feishu-docget/
├── README.md
├── feishu-docget.properties     # 本地配置，包含飞书密钥，已被 .gitignore 忽略
├── run.bat                      # Windows 启动脚本
├── run.sh                       # Linux/macOS 启动脚本
├── src/
│   ├── app.py                   # Flask Web、API、任务队列和管理后台入口
│   ├── cli/feishu2word.py       # 命令行入口
│   ├── core/
│   │   ├── config_loader.py     # 配置加载、默认配置补全、日志初始化
│   │   ├── feishu_client.py     # 飞书 Token、文档块、媒体/画板下载
│   │   ├── image_processor.py   # 图片裁剪
│   │   ├── stats.py             # 下载统计
│   │   └── utils.py             # 通用工具
│   ├── services/doc_service.py  # 单文档处理编排
│   ├── converters/docx/
│   │   ├── converter.py         # 飞书 Block -> Word 内容树
│   │   ├── cleaner.py           # 模板复制、样式清洗、图片/表格/代码块后处理
│   │   └── style_manager.py     # 6 种表格样式
│   └── web/templates/
│       ├── index.html           # 前台下载页面
│       ├── dashboard.html       # 管理后台
│       └── login.html           # 管理后台登录页
├── template/                    # Word 模板和同名预览图片
├── output/                      # 导出结果
└── logs/                        # 运行日志和下载统计
```

## 快速开始

### 1. 准备飞书应用

在飞书开放平台创建自建应用，获取：

- `App ID`
- `App Secret`

建议开通云文档、电子表格、知识库、素材下载、画板下载等与导出范围相关的权限。对于企业内部文档，还需要在文档页面把机器人添加为协作者，否则接口会返回无权限。

### 2. 配置本地文件

首次启动会自动生成 `feishu-docget.properties`。也可以手动创建，至少需要配置：

```properties
feishu.app_id=cli_xxxxxxxxxxxxxxxx
feishu.app_secret=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
server.port=7800
admin.path=/admin
admin.password=change-me
template.default=Hawkon.docx
workspace.dir=.
template.dir=template
output.dir=output
log.dir=logs
```

注意：`feishu-docget.properties` 包含敏感凭据，不要提交到 Git。

### 3. 启动 Web 服务

Windows：

```bat
run.bat
```

Linux/macOS：

```bash
chmod +x run.sh
./run.sh
```

启动后访问：

```text
http://127.0.0.1:7800/
```

管理后台默认路径由 `admin.path` 控制，默认：

```text
http://127.0.0.1:7800/admin
```

### 4. 命令行导出

```bash
python src/cli/feishu2word.py "https://example.feishu.cn/wiki/xxxx" \
  --template Hawkon.docx \
  --style 3 \
  --cover
```

参数说明：

- `url`：飞书文档链接。
- `--template/-t`：模板文件名，默认读取 `template.default`。
- `--style/-s`：表格样式 ID，范围 `1-6`。
- `--output/-o`：输出目录，默认读取 `output.dir`。
- `--cover/-c`：从模板追加封面。

## 使用说明

### 前台下载

1. 在左侧选择 Word 模板。
2. 选择表格样式。
3. 输入飞书文档链接。
4. 按需打开“高级选项”调整清洗规则。
5. 点击“确认下载”，任务会进入队列并显示进度。

导出结果默认写入：

```text
output/<doc_id>/<文档标题>.docx
```

图片资源会放入同级 `img/` 目录，重复下载同一文档时会尽量复用历史图片资源。

### 高级选项

高级选项按模块拆分：

- 基础设置：封面、标题、页边距、文档信息。
- 文本设置：文本规则、列表样式、正文字号、段落间距。
- 图片设置：图片尺寸、图片对齐、表格图片覆盖。
- 表格设置：表格规则、布局、内容对齐、段落间距、边框管理。
- 代码块：代码外观、布局、段落间距、缩进规则、边框设置、表格代码块覆盖。

高级选项只影响当前下载任务，不会自动修改模板文件。

### 模板管理

模板目录由 `template.dir` 控制，默认是 `template/`。

一个完整模板通常包含：

```text
template/模板名.docx
template/模板名.png
```

`.docx` 用于 Word 样式、页眉页脚和封面；同名 `.png` 用于前台预览。管理后台支持上传、替换、重命名、删除、设为默认模板。

首页也支持模板上传。如果配置了：

```properties
template.password.long_term=
template.password.one_time=
```

则非管理员上传时会校验对应密码；留空时非管理员上传会被拒绝。已登录管理后台时会跳过模板上传密码校验，并按长期保存模式处理。

### 管理后台

后台功能包括：

- 项目管理：查看、下载、删除已导出的项目。
- 配置管理：在线编辑 `feishu-docget.properties` 中的配置项。
- 下载统计：查看任务记录，支持批量删除统计记录。
- 日志检查：查看、刷新、删除日志文件。
- 模板维护：上传、预览、重命名、设为默认、删除模板。
- 系统管理：封装部分系统脚本操作。

后台登录密码由 `admin.password` 控制。

## 表格样式

前台提供 6 种表格样式：

1. 深蓝表头 + 白字加粗
2. 浅蓝表头 + 网格边框
3. 浅灰表头 + 细网格边框
4. 全黑实线
5. 上下黑边 + 中间灰竖线
6. 黑表头 + 斑马纹

这些样式在 `src/converters/docx/style_manager.py` 中维护，前台预览 CSS 和 Word 写入逻辑保持同一套语义。

如需覆盖样式边框，可在“高级选项 -> 表格设置 -> 边框管理”中启用自定义表格边框。启用后会对普通表格单元格写入统一颜色和上下左右边框配置。

## 转换流程

整体流程如下：

1. `src/app.py` 接收 Web/API 请求，创建任务并放入下载队列。
2. `src/services/doc_service.py` 根据链接解析文档 ID，创建输出目录。
3. `FeishuClient` 获取飞书文档元信息、文档块、图片和画板资源。
4. `FeishuDocxConverter` 将飞书 Block 递归渲染为 Word 内容。
5. `TableStyleManager` 应用前台选择的表格样式。
6. `clean_document` 复制模板样式、页眉页脚、封面，并执行图片、表格、代码块、正文、页边距等清洗。
7. `apply_document_info` 写入 Word 元数据。
8. Web 前台轮询任务状态，任务完成后提供下载。

## 配置项

常用配置：

```properties
# 飞书应用
feishu.app_id=
feishu.app_secret=

# 服务
server.port=7800
admin.path=/admin
admin.password=change-me

# 模板
template.default=Hawkon.docx
template.dir=template
template.password.long_term=
template.password.one_time=

# 导出
image.max_width=16
image.max_height=23
download.threads=4
max.concurrent.downloads=1
download_images=True

# 路径和日志
workspace.dir=.
output.dir=output
output.max_size=10G
log.dir=logs
log.level=INFO
log.max_size=20M
```

`ConfigLoader` 会在启动时补齐缺失配置项，并创建日志、输出和模板目录。

## 常见问题

### 提示无权限或 Permission Denied

请检查：

- 飞书应用是否开通对应云文档权限。
- 文档是否已添加机器人为协作者。
- 链接是否属于当前企业空间，且应用有访问范围。

### 图片、画板没有出现在 Word 中

请检查：

- `download_images=True`。
- 应用是否具备素材下载权限。
- 日志中是否有媒体下载 403 或超时。

### 生成的 Word 样式和模板不一致

高级选项中的清洗规则优先级高于模板。若希望完全跟随模板，可把相关高级设置留空或关闭，例如正文样式、图片限制、表格布局、边框管理等。

### 表格边框和前台预览不一致

优先确认是否启用了“自定义表格边框”。未启用时使用 6 种预设表格样式；启用后自定义边框会覆盖普通表格单元格的边框。

### 输出目录越来越大

`output.max_size` 会用于输出目录清理逻辑。也可以在管理后台删除旧项目，或下载全部文件后手动归档。

## 开发提示

- Web 入口：`src/app.py`
- 前台页面：`src/web/templates/index.html`
- 管理后台：`src/web/templates/dashboard.html`
- 转换主逻辑：`src/converters/docx/converter.py`
- 格式清洗：`src/converters/docx/cleaner.py`
- 表格样式：`src/converters/docx/style_manager.py`
- 飞书接口：`src/core/feishu_client.py`

开发时建议先运行：

```bash
python -m compileall ./src
```

修改前台后，刷新 `http://127.0.0.1:7800/` 验证高级选项、模板选择、表格预览和任务队列。

## 安全说明

- 不要提交 `feishu-docget.properties`、日志、导出文件和私有模板。
- 管理后台暴露了下载、删除、配置和系统操作能力，部署到公网前务必设置强密码，并放在可信网络或反向代理鉴权之后。
- 本项目仅供学习、归档和内部自动化场景使用，请遵守飞书平台规则和所在组织的数据合规要求。

## License

本项目使用 `LICENSE` 文件中的开源许可。
