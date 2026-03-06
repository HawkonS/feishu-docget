#!/usr/bin/env python3
import sys
import os
import argparse
import time

current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(os.path.dirname(current_dir))
sys.path.insert(0, project_root)

try:
    from src.services.doc_service import process_document
    from src.core.config_loader import config, ConfigLoader
except ImportError as e:
    print(f"导入模块错误: {e}")
    print("请在项目根目录下运行此脚本，或确保 Python 路径正确。")
    sys.exit(1)

def main():
    parser = argparse.ArgumentParser(description='飞书文档下载工具命令行版')
    parser.add_argument('url', help='飞书文档链接')
    parser.add_argument('--template', '-t', help='模板文件名称 (例如: Hawkon.docx)', default=None)
    parser.add_argument('--style', '-s', help='表格样式 ID (1-6)', default=None)
    parser.add_argument('--output', '-o', help='输出目录 (默认使用配置中的 output.dir)', default=None)
    parser.add_argument('--cover', '-c', action='store_true', help='是否添加封面')
    
    args = parser.parse_args()
    
    # 确保配置已加载
    if not config:
        print("错误: 配置未加载。")
        sys.exit(1)

    # 准备参数
    base_dir = os.path.abspath(config.get('workspace.dir', '.'))
    
    # 确定模板路径
    template_name = args.template
    if not template_name:
        template_name = config.get('template.default', 'Hawkon.docx')
    
    if template_name and not template_name.lower().endswith('.docx'):
        template_name += '.docx'
        
    template_dir = os.path.join(base_dir, config.get('template.dir', 'template'))
    template_path = os.path.join(template_dir, template_name)
    
    if not os.path.exists(template_path):
        print(f"警告: 未找到模板文件于 {template_path}")
        
    # 确定输出目录
    if args.output:
        output_root = args.output
    else:
        output_root = os.path.join(base_dir, config.get('output.dir', 'output'))

    print("=" * 50)
    print(f"开始下载任务")
    print(f"文档链接: {args.url}")
    print(f"使用模板: {template_name}")
    print(f"表格样式: {args.style if args.style else '默认'}")
    print(f"输出目录: {output_root}")
    print("=" * 50)

    # 进度回调
    def progress_callback(progress, message, type='info'):
        # 简单的进度条显示
        bar_length = 20
        filled_length = int(bar_length * progress // 100)
        bar = '█' * filled_length + '-' * (bar_length - filled_length)
        print(f'\r[{bar}] {progress}% {message}', end='')
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
            add_cover=args.cover
        )
        end_time = time.time()
        duration = end_time - start_time
        
        print("\n" + "=" * 50)
        print(f"下载成功！")
        print(f"耗时: {duration:.2f} 秒")
        print(f"保存路径: {result['docx_path']}")
        print("=" * 50)
        
    except KeyboardInterrupt:
        print("\n任务已取消")
        sys.exit(1)
    except Exception as e:
        print(f"\n任务失败: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()
