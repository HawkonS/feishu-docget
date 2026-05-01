#!/usr/bin/env python3
import argparse
import json
import os
import re
import sys
import time
from datetime import datetime

current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(os.path.dirname(current_dir))
sys.path.insert(0, project_root)

try:
    from src.services.doc_service import process_document
    from src.core.config_loader import config
    from src.converters.docx.style_manager import TableStyleManager
except ImportError as e:
    print(f"导入模块错误: {e}")
    print("请在项目根目录下运行此脚本，或确保 Python 路径正确。")
    sys.exit(1)


BORDER_TYPES = ("single", "none", "double", "dotted", "dashed")
ALIGN_TYPES = ("none", "center", "left", "right")
TEXT_UNITS = ("lines", "pt")
MARGIN_PRESETS = {
    "normal": {"top": 2.54, "bottom": 2.54, "left": 3.18, "right": 3.18},
    "narrow": {"top": 1.27, "bottom": 1.27, "left": 1.27, "right": 1.27},
    "wide": {"top": 2.54, "bottom": 2.54, "left": 5.08, "right": 5.08},
    "moderate": {"top": 2.54, "bottom": 2.54, "left": 1.91, "right": 1.91},
}
DATETIME_FORMATS = (
    "%Y-%m-%dT%H:%M",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M",
    "%Y/%m/%d %H:%M:%S",
)


def non_negative_float(value):
    try:
        parsed = float(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("必须是数字") from exc
    if parsed < 0:
        raise argparse.ArgumentTypeError("必须大于或等于 0")
    return parsed


def non_negative_int(value):
    try:
        parsed = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("必须是整数") from exc
    if parsed < 0:
        raise argparse.ArgumentTypeError("必须大于或等于 0")
    return parsed


def bounded_int(min_value, max_value):
    def parser(value):
        parsed = non_negative_int(value)
        if parsed < min_value or parsed > max_value:
            raise argparse.ArgumentTypeError(f"范围必须在 {min_value}-{max_value} 之间")
        return parsed
    return parser


def color_value(value):
    color = str(value or "").strip()
    if re.fullmatch(r"#?[0-9a-fA-F]{6}", color):
        return color if color.startswith("#") else f"#{color}"
    raise argparse.ArgumentTypeError("颜色必须是 #RRGGBB 或 RRGGBB")


def current_local_minute():
    return datetime.now().strftime("%Y-%m-%dT%H:%M")


def normalize_datetime(value):
    if value is None:
        return ""
    raw = str(value).strip()
    if not raw:
        return ""
    if raw.lower() == "now":
        return current_local_minute()
    return raw


def validate_document_info(document_info):
    if not document_info:
        return None
    field_labels = {"created": "创建时间", "modified": "上次修改时间", "lastPrinted": "上次打印时间"}
    for key, label in field_labels.items():
        raw_value = str(document_info.get(key) or "").strip()
        if not raw_value:
            continue
        if any(_is_datetime(raw_value, fmt) for fmt in DATETIME_FORMATS):
            continue
        return f"{label}格式无效，请使用 YYYY-MM-DDTHH:MM、YYYY-MM-DD HH:MM 或 now"
    return None


def _is_datetime(value, fmt):
    try:
        datetime.strptime(value, fmt)
        return True
    except ValueError:
        return False


def list_templates(base_dir):
    template_dir = os.path.join(base_dir, config.get("template.dir", "template"))
    if not os.path.isdir(template_dir):
        return []
    default_template = config.get("template.default", "template.docx")
    templates = []
    for name in sorted(os.listdir(template_dir)):
        if name.startswith("temp_") or not name.lower().endswith(".docx"):
            continue
        path = os.path.join(template_dir, name)
        templates.append({
            "name": name,
            "path": path,
            "size": os.path.getsize(path) if os.path.exists(path) else 0,
            "default": name == default_template,
        })
    return templates


def print_templates(base_dir):
    templates = list_templates(base_dir)
    if not templates:
        print("未找到模板文件")
        return
    for item in templates:
        suffix = " (默认)" if item["default"] else ""
        print(f"- {item['name']}{suffix}  {item['size']} bytes")


def print_styles():
    for item in TableStyleManager.list_styles():
        print(f"- {item['id']}: {item['name']}")


def build_body_style(args):
    if not any(v is not None for v in (
        args.body_font_size,
        args.body_line_spacing,
        args.body_space_before,
        args.body_space_after,
    )):
        return None
    return {
        "fontSize": args.body_font_size,
        "lineSpacing": args.body_line_spacing,
        "lineSpacingUnit": args.body_line_spacing_unit,
        "spaceBefore": args.body_space_before,
        "spaceBeforeUnit": args.body_space_before_unit,
        "spaceAfter": args.body_space_after,
        "spaceAfterUnit": args.body_space_after_unit,
    }


def build_image_style(args):
    image_style = None
    if args.image_max_width is not None or args.image_max_height is not None or args.image_align != "none":
        image_style = {
            "maxWidth": args.image_max_width,
            "maxHeight": args.image_max_height,
            "align": None if args.image_align == "none" else args.image_align,
        }
    if args.table_image_max_width is not None or args.table_image_max_height is not None:
        if image_style is None:
            image_style = {}
        image_style["tableImageStyle"] = {
            "maxWidth": args.table_image_max_width,
            "maxHeight": args.table_image_max_height,
        }
    return image_style


def build_border_config(args, prefix):
    return {
        "top": {
            "type": getattr(args, f"{prefix}_border_top_type"),
            "width": getattr(args, f"{prefix}_border_top_width"),
        },
        "bottom": {
            "type": getattr(args, f"{prefix}_border_bottom_type"),
            "width": getattr(args, f"{prefix}_border_bottom_width"),
        },
        "left": {
            "type": getattr(args, f"{prefix}_border_left_type"),
            "width": getattr(args, f"{prefix}_border_left_width"),
        },
        "right": {
            "type": getattr(args, f"{prefix}_border_right_type"),
            "width": getattr(args, f"{prefix}_border_right_width"),
        },
    }


def build_table_config(args):
    return {
        "forceClearIndent": args.table_force_clear_indent,
        "forceClearImageSpace": args.table_force_clear_image_space,
        "autoFit": args.table_auto_fit,
        "width": args.table_width,
        "minColWidth": args.table_min_col_width,
        "headerAlign": args.table_header_align,
        "contentAlign": args.table_content_align,
        "contentImageAlign": args.table_content_image_align,
        "lineSpacing": args.table_line_spacing,
        "spaceBefore": args.table_space_before,
        "spaceAfter": args.table_space_after,
        "borderEnabled": args.table_border,
        "borderColor": args.table_border_color,
        "borders": build_border_config(args, "table") if args.table_border else None,
    }


def build_margin_config(args):
    if args.margin_preset in MARGIN_PRESETS:
        return dict(MARGIN_PRESETS[args.margin_preset])
    margin_values = {
        "top": args.margin_top,
        "bottom": args.margin_bottom,
        "left": args.margin_left,
        "right": args.margin_right,
    }
    margin_config = {key: value for key, value in margin_values.items() if value is not None}
    return margin_config or None


def build_code_block_config(args):
    return {
        "bgColor": args.code_bg_color,
        "fontColor": args.code_font_color,
        "fontFamily": args.code_font_family,
        "fontSize": args.code_font_size,
        "align": args.code_align,
        "tableWidth": args.code_table_width,
        "innerTableWidth": args.code_inner_table_width,
        "lineSpacing": args.code_line_spacing,
        "spaceBefore": args.code_space_before,
        "spaceAfter": args.code_space_after,
        "forceClearIndent": args.code_force_clear_indent,
        "borderColor": args.code_border_color,
        "borders": build_border_config(args, "code"),
    }


def build_document_info(args):
    document_info = {
        "author": args.doc_author or "",
        "lastModifiedBy": args.doc_last_modified_by or "",
        "created": normalize_datetime(args.doc_created),
        "modified": normalize_datetime(args.doc_modified),
        "lastPrinted": normalize_datetime(args.doc_last_printed),
        "totalTime": args.doc_total_time,
        "title": args.doc_title or "",
        "category": args.doc_category or "",
        "subject": args.doc_subject or "",
        "company": args.doc_company or "",
        "template": args.doc_template or "",
    }
    if any(value not in ("", None) for value in document_info.values()):
        return document_info
    return None


def build_effective_options(args):
    return {
        "addCover": args.cover,
        "addTitle": args.add_title,
        "ignoreMention": args.ignore_mention,
        "ignoreTemplateHeadingNum": args.ignore_template_heading_num,
        "unorderedListStyle": args.unordered_list_style,
        "bodyStyle": build_body_style(args),
        "imageStyle": build_image_style(args),
        "tableConfig": build_table_config(args),
        "marginConfig": build_margin_config(args),
        "codeBlockConfig": build_code_block_config(args),
        "documentInfo": build_document_info(args),
    }


def add_border_arguments(group, prefix, title, default_width):
    group.add_argument(f"--{prefix}-border-color", type=color_value, default="#D9D9D9", help=f"{title}统一边框颜色")
    for edge, edge_name in (("top", "上"), ("bottom", "下"), ("left", "左"), ("right", "右")):
        group.add_argument(
            f"--{prefix}-border-{edge}-type",
            choices=BORDER_TYPES,
            default="single",
            help=f"{title}{edge_name}边框线型",
        )
        group.add_argument(
            f"--{prefix}-border-{edge}-width",
            type=non_negative_int,
            default=default_width,
            help=f"{title}{edge_name}边框粗细，单位为 Word eighth-point",
        )


def create_parser():
    parser = argparse.ArgumentParser(
        description="飞书文档下载工具命令行版，支持前台下载页和高级选项的同等导出能力。",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("url", nargs="?", help="飞书文档链接")

    core = parser.add_argument_group("基础下载")
    template_choice = core.add_mutually_exclusive_group()
    template_choice.add_argument("--template", "-t", help="模板文件名称，如 Hawkon.docx；默认读取 template.default")
    template_choice.add_argument("--template-path", help="直接使用本地 .docx 模板路径，不要求位于 template.dir")
    core.add_argument("--style", "-s", choices=[str(i) for i in range(1, 7)], help="表格样式 ID，范围 1-6")
    core.add_argument("--output", "-o", help="输出目录；默认读取 output.dir")
    core.add_argument("--list-templates", action="store_true", help="列出 template.dir 下可用模板后退出")
    core.add_argument("--list-styles", action="store_true", help="列出前台 6 种表格样式后退出")
    core.add_argument("--print-options", action="store_true", help="打印本次 CLI 组装出的前台兼容高级选项 JSON 后退出")

    basic = parser.add_argument_group("高级选项 - 基础设置")
    basic.add_argument("--cover", "-c", action=argparse.BooleanOptionalAction, default=True, help="添加模板封面/首页")
    basic.add_argument("--add-title", action=argparse.BooleanOptionalAction, default=False, help="在正文内容中添加标题")
    basic.add_argument("--margin-preset", choices=("none", "custom", "normal", "narrow", "wide", "moderate"), default="none", help="页边距预设")
    basic.add_argument("--margin-top", type=non_negative_float, help="上边距，单位 cm")
    basic.add_argument("--margin-bottom", type=non_negative_float, help="下边距，单位 cm")
    basic.add_argument("--margin-left", type=non_negative_float, help="左边距，单位 cm")
    basic.add_argument("--margin-right", type=non_negative_float, help="右边距，单位 cm")
    basic.add_argument("--doc-author", help="Word 文档信息：作者")
    basic.add_argument("--doc-last-modified-by", help="Word 文档信息：上次修改者")
    basic.add_argument("--doc-created", help="Word 文档信息：创建时间，支持 YYYY-MM-DDTHH:MM 或 now")
    basic.add_argument("--doc-modified", help="Word 文档信息：上次修改时间，支持 YYYY-MM-DDTHH:MM 或 now")
    basic.add_argument("--doc-last-printed", help="Word 文档信息：上次打印时间，支持 YYYY-MM-DDTHH:MM 或 now")
    basic.add_argument("--doc-total-time", type=non_negative_int, help="Word 文档信息：编辑时间统计，单位分钟")
    basic.add_argument("--doc-title", help="Word 文档信息：标题")
    basic.add_argument("--doc-category", help="Word 文档信息：类别")
    basic.add_argument("--doc-subject", help="Word 文档信息：主题")
    basic.add_argument("--doc-company", help="Word 文档信息：单位")
    basic.add_argument("--doc-template", help="Word 文档信息：模板")

    text = parser.add_argument_group("高级选项 - 文本设置")
    text.add_argument("--ignore-mention", action=argparse.BooleanOptionalAction, default=False, help="忽略文本中 @ 人员转换")
    text.add_argument("--ignore-template-heading-num", action=argparse.BooleanOptionalAction, default=False, help="强制移除模板标题自动编号")
    text.add_argument("--unordered-list-style", choices=("none", "default", "square", "diamond", "arrow"), default="default", help="无序列表样式")
    text.add_argument("--body-font-size", type=non_negative_float, help="正文字号，单位磅")
    text.add_argument("--body-line-spacing", type=non_negative_float, help="正文行间距")
    text.add_argument("--body-line-spacing-unit", choices=TEXT_UNITS, default="lines", help="正文行间距单位")
    text.add_argument("--body-space-before", type=non_negative_float, help="正文段前")
    text.add_argument("--body-space-before-unit", choices=TEXT_UNITS, default="lines", help="正文段前单位")
    text.add_argument("--body-space-after", type=non_negative_float, help="正文段后")
    text.add_argument("--body-space-after-unit", choices=TEXT_UNITS, default="lines", help="正文段后单位")

    image = parser.add_argument_group("高级选项 - 图片设置")
    image.add_argument("--image-max-width", type=non_negative_float, help="普通图片最大宽度，单位 cm")
    image.add_argument("--image-max-height", type=non_negative_float, help="普通图片最大高度，单位 cm")
    image.add_argument("--image-align", choices=ALIGN_TYPES, default="center", help="普通图片对齐方式")

    table = parser.add_argument_group("高级选项 - 表格设置")
    table.add_argument("--table-force-clear-indent", action=argparse.BooleanOptionalAction, default=True, help="强制清除表格缩进")
    table.add_argument("--table-force-clear-image-space", action=argparse.BooleanOptionalAction, default=True, help="强制清除表格图片段前间距")
    table.add_argument("--table-auto-fit", action=argparse.BooleanOptionalAction, default=True, help="启用表格自适应")
    table.add_argument("--table-width", default="100%", help="表格宽度，如 100%% 或 15cm")
    table.add_argument("--table-min-col-width", type=bounded_int(1, 100), default=8, help="表格单列最小宽度，单位字符")
    table.add_argument("--table-header-align", choices=("center", "left", "right"), default="center", help="表头文字对齐")
    table.add_argument("--table-content-align", choices=("center", "left", "right"), default="left", help="内容文字对齐")
    table.add_argument("--table-content-image-align", choices=("center", "left", "right"), default="left", help="内容图片对齐")
    table.add_argument("--table-line-spacing", type=non_negative_float, help="表格内容行间距，单位行")
    table.add_argument("--table-space-before", type=non_negative_float, help="表格内容段前距，单位行")
    table.add_argument("--table-space-after", type=non_negative_float, help="表格内容段后距，单位行")
    table.add_argument("--code-inner-table-width", type=non_negative_float, help="表格内代码块宽度，单位 cm")
    table.add_argument("--table-image-max-width", type=non_negative_float, help="表格内图片最大宽度，单位 cm")
    table.add_argument("--table-image-max-height", type=non_negative_float, help="表格内图片最大高度，单位 cm")
    table.add_argument("--table-border", action=argparse.BooleanOptionalAction, default=False, help="启用自定义表格边框")
    add_border_arguments(table, "table", "表格", 6)

    code = parser.add_argument_group("高级选项 - 代码块")
    code.add_argument("--code-bg-color", type=color_value, default="#F5F5F5", help="代码块背景颜色")
    code.add_argument("--code-font-color", type=color_value, default="#000000", help="代码块字体颜色")
    code.add_argument("--code-font-family", default="Courier New", help="代码块字体")
    code.add_argument("--code-font-size", type=non_negative_float, default=9, help="代码块字号，单位磅")
    code.add_argument("--code-align", choices=("left", "center", "right"), default="left", help="代码块对齐方式")
    code.add_argument("--code-table-width", type=non_negative_float, help="代码块表格宽度，单位 cm")
    code.add_argument("--code-line-spacing", type=non_negative_float, help="代码块行间距，单位行")
    code.add_argument("--code-space-before", type=non_negative_float, help="代码块段前距，单位行")
    code.add_argument("--code-space-after", type=non_negative_float, help="代码块段后距，单位行")
    code.add_argument("--code-force-clear-indent", action=argparse.BooleanOptionalAction, default=True, help="强制删除代码块缩进")
    add_border_arguments(code, "code", "代码块", 4)

    return parser


def resolve_template_path(args, base_dir):
    if args.template_path:
        template_path = os.path.abspath(args.template_path)
        template_name = os.path.basename(template_path)
        return template_name, template_path

    template_name = args.template or config.get("template.default", "Hawkon.docx")
    if template_name and not template_name.lower().endswith(".docx"):
        template_name += ".docx"
    template_dir = os.path.join(base_dir, config.get("template.dir", "template"))
    return template_name, os.path.join(template_dir, template_name) if template_name else ""


def print_effective_options(args, template_name, template_path, output_root, effective_options):
    payload = {
        "url": args.url,
        "template": template_name,
        "templatePath": template_path,
        "tableStyle": args.style,
        "output": output_root,
        **effective_options,
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))


def main():
    parser = create_parser()
    args = parser.parse_args()

    if not config:
        print("错误: 配置未加载。")
        sys.exit(1)

    base_dir = os.path.abspath(config.get("workspace.dir", "."))

    if args.list_templates:
        print_templates(base_dir)
    if args.list_styles:
        print_styles()
    if (args.list_templates or args.list_styles) and not args.url and not args.print_options:
        return

    if not args.url:
        parser.error("缺少飞书文档链接；仅查看模板或样式时可使用 --list-templates / --list-styles")

    template_name, template_path = resolve_template_path(args, base_dir)
    if template_path and not os.path.exists(template_path):
        print(f"警告: 未找到模板文件于 {template_path}")

    output_root = args.output or os.path.join(base_dir, config.get("output.dir", "output"))
    effective_options = build_effective_options(args)
    document_info_error = validate_document_info(effective_options.get("documentInfo"))
    if document_info_error:
        parser.error(document_info_error)

    if args.print_options:
        print_effective_options(args, template_name, template_path, output_root, effective_options)
        return

    print("=" * 50)
    print("开始下载任务")
    print(f"文档链接: {args.url}")
    print(f"使用模板: {template_name}")
    print(f"表格样式: {args.style if args.style else '默认'}")
    print(f"输出目录: {output_root}")
    print(f"添加封面: {'是' if effective_options['addCover'] else '否'}")
    print("=" * 50)

    def progress_callback(progress, message, type="info"):
        bar_length = 20
        filled_length = int(bar_length * progress // 100)
        bar = "█" * filled_length + "-" * (bar_length - filled_length)
        print(f"\r[{bar}] {progress}% {message}", end="")
        if progress >= 100:
            print()

    try:
        start_time = time.time()
        result = process_document(
            doc_url=args.url,
            template_path=template_path,
            table_style=args.style,
            base_dir=base_dir,
            output_root=output_root,
            progress_cb=progress_callback,
            add_cover=effective_options["addCover"],
            unordered_list_style=effective_options["unorderedListStyle"],
            body_style=effective_options["bodyStyle"],
            image_style=effective_options["imageStyle"],
            ignore_mention=effective_options["ignoreMention"],
            ignore_template_heading_num=effective_options["ignoreTemplateHeadingNum"],
            table_config=effective_options["tableConfig"],
            margin_config=effective_options["marginConfig"],
            code_block_config=effective_options["codeBlockConfig"],
            document_info=effective_options["documentInfo"],
            add_title=effective_options["addTitle"],
        )
        duration = time.time() - start_time

        print("\n" + "=" * 50)
        print("下载成功！")
        print(f"耗时: {duration:.2f} 秒")
        print(f"保存路径: {result['docx_path']}")
        print("=" * 50)

    except KeyboardInterrupt:
        print("\n任务已取消")
        sys.exit(1)
    except Exception as e:
        print(f"\n任务失败: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
