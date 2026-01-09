#src\codecollate\cli.py
import argparse
import logging
import sys
import os
from pathlib import Path

from .collator import SourceCodeCollator

def setup_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(levelname)s: %(message)s"
    )

def _prompt_input(prompt: str) -> str:
    return input(prompt).strip()

def _prompt_confirm(prompt: str, default: bool = True) -> bool:
    hint = "Y/n" if default else "y/N"
    while True:
        value = input(f"{prompt} ({hint}): ").strip().lower()
        if not value:
            return default
        if value in {"y", "yes"}:
            return True
        if value in {"n", "no"}:
            return False
        print("请输入 y 或 n。")

def main() -> None:
    """命令行工具的主入口函数 (交互式, 支持配置)。"""
    parser = argparse.ArgumentParser(
        prog="codecollate",
        description="一个专业的源代码整理工具，可将源代码集合格式化为符合特定规范的DOCX文档。",
        formatter_class=argparse.RawTextHelpFormatter
    )
    
    parser.add_argument("source_dir", type=str, nargs='?', default=None, help="【可选】包含源代码文件的目录路径。")
    parser.add_argument("software_name", type=str, nargs='?', default=None, help="【可选】软件的名称。")
    parser.add_argument("version", type=str, nargs='?', default=None, help="【可选】软件的版本号。")
    parser.add_argument("-o", "--output", type=str, default=".", help="指定输出文档的目录路径。\n(默认值: 工作目录)")
    parser.add_argument("--config", type=str, default=None, help="指定一个自定义YAML配置文件的路径。\n如果未提供，将自动查找当前目录下的 'codecollate_config.yaml'。")
    parser.add_argument("-w", "--workdir", type=str, default=None, help="设置工作目录（程序会在该目录下查找源代码与配置文件）。")
    parser.add_argument("-v", "--verbose", action="store_true", help="启用详细日志输出模式，用于调试。")

    args = parser.parse_args()
    
    setup_logging(args.verbose)
    print("CodeCollate-CLI-P - 专业源代码文档整理工具")

    try:
        if not args.workdir:
            while True:
                workdir_str = _prompt_input("请输入工作目录路径: ")
                workdir_path = Path(workdir_str).resolve()
                if workdir_path.is_dir():
                    args.workdir = str(workdir_path)
                    break
                else:
                    print(f"错误：工作目录 '{workdir_path}' 不存在或不是一个目录，请重新输入。")
        else:
            workdir_path = Path(args.workdir).resolve()
            if not workdir_path.is_dir():
                print(f"错误：命令行提供的工作目录 '{workdir_path}' 不存在或不是一个目录。")
                sys.exit(1)

        os.chdir(workdir_path)

        if not args.source_dir:
            source_path = workdir_path.resolve()
            args.source_dir = str(source_path)
        else:
            input_path = Path(args.source_dir)
            source_path = (input_path if input_path.is_absolute() else workdir_path / input_path).resolve()
            if not source_path.is_dir():
                print(f"错误：命令行提供的路径 '{source_path}' 不存在或不是一个目录。")
                sys.exit(1)

        if not args.software_name:
            args.software_name = _prompt_input("请输入软件名称: ")
        if not args.version:
            args.version = _prompt_input("请输入版本号: ")

        output_input_path = Path(args.output)
        output_path = (output_input_path if output_input_path.is_absolute() else workdir_path / output_input_path).resolve()
        args.output = str(output_path)

        print("\n请确认以下信息：")
        print(f"工作目录: {workdir_path}")
        print(f"源代码目录: {args.source_dir}")
        print(f"软件名称: {args.software_name}")
        print(f"版本号: {args.version}")
        print(f"输出目录: {args.output}")
        if args.config:
            print(f"配置文件: {args.config}")

        if not _prompt_confirm("是否开始执行整理任务？", default=True):
            print("操作已取消。")
            sys.exit(0)

        collator = SourceCodeCollator(
            source_dir=Path(args.source_dir),
            software_name=args.software_name,
            version=args.version,
            output_dir=output_path,
            config_path=args.config
        )
        
        docx_file = None

        print("正在整理源代码，请稍候...")
        docx_file = collator.run()

        if docx_file:
            print("\n" + "="*50)
            print("源代码整理成功！")
            print(f"Word 文档已生成: {docx_file.resolve()}")
            print("="*50)
            
    except Exception as e:
        logging.error("处理失败，以下是详细信息：", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()
