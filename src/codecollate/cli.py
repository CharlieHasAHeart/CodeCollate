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
        format="[%(levelname)s] %(message)s"
    )

def _prompt_input(prompt: str) -> str:
    return input(f"[INPUT] {prompt}").strip()

def _prompt_confirm(prompt: str, default: bool = True) -> bool:
    hint = "Y/n" if default else "y/N"
    while True:
        value = input(f"[CONFIRM] {prompt} ({hint}): ").strip().lower()
        if not value:
            return default
        if value in {"y", "yes"}:
            return True
        if value in {"n", "no"}:
            return False
        print("[ERROR] Please enter y or n.")

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
    print("[INFO] CodeCollate-CLI-P - Professional Source Code Document Tool")

    try:
        if not args.workdir:
            while True:
                workdir_str = _prompt_input("Enter working directory path: ")
                workdir_path = Path(workdir_str).resolve()
                if workdir_path.is_dir():
                    args.workdir = str(workdir_path)
                    break
                else:
                    print(f"[ERROR] Working directory '{workdir_path}' does not exist or is not a directory.")
        else:
            workdir_path = Path(args.workdir).resolve()
            if not workdir_path.is_dir():
                print(f"[ERROR] Provided working directory '{workdir_path}' does not exist or is not a directory.")
                sys.exit(1)

        os.chdir(workdir_path)

        if not args.source_dir:
            source_path = workdir_path.resolve()
            args.source_dir = str(source_path)
        else:
            input_path = Path(args.source_dir)
            source_path = (input_path if input_path.is_absolute() else workdir_path / input_path).resolve()
            if not source_path.is_dir():
                print(f"[ERROR] Provided source directory '{source_path}' does not exist or is not a directory.")
                sys.exit(1)

        if not args.software_name:
            args.software_name = _prompt_input("Enter software name: ")
        if not args.version:
            args.version = _prompt_input("Enter version: ")

        output_input_path = Path(args.output)
        output_path = (output_input_path if output_input_path.is_absolute() else workdir_path / output_input_path).resolve()
        args.output = str(output_path)

        print("\n[INFO] Please confirm the following:")
        print(f"[INFO] Working directory: {workdir_path}")
        print(f"[INFO] Source directory: {args.source_dir}")
        print(f"[INFO] Software name: {args.software_name}")
        print(f"[INFO] Version: {args.version}")
        print(f"[INFO] Output directory: {args.output}")
        if args.config:
            print(f"[INFO] Config file: {args.config}")

        if not _prompt_confirm("Start the collation task?", default=True):
            print("[INFO] Operation cancelled.")
            sys.exit(0)

        collator = SourceCodeCollator(
            source_dir=Path(args.source_dir),
            software_name=args.software_name,
            version=args.version,
            output_dir=output_path,
            config_path=args.config
        )
        
        docx_file = None

        print("[INFO] Processing source code, please wait...")
        docx_file = collator.run()

        if docx_file:
            print("\n" + "="*50)
            print("[SUCCESS] Collation completed.")
            print(f"[OUTPUT] Word document generated: {docx_file.resolve()}")
            print("="*50)
            
    except Exception as e:
        logging.error("处理失败，以下是详细信息：", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()
