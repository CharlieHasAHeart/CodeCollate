# CodeCollate-CLI-P

[![Python Version](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

`CodeCollate-CLI-P` 是一个专业的命令行工具，旨在将源代码集合自动化地整理、格式化为符合特定规范（如软件著作权申请）的专业文档。它能智能地处理代码，扩充内容以满足页数要求，并最终生成高质量的 DOCX 文件。

## ✨ 核心功能

- **自动化整理**: 递归扫描指定目录下的源代码文件（默认包含 `.html`, `.js`, `.css`, `.py`, `.java`, `.ts`, `.tsx`）。
- **智能清洗**: 自动移除注释和空行（可在配置中关闭）。
- **内容扩充**: 当源代码行数不足时，采用“受控重复与代码块标注”策略扩充至目标页数（默认约 63 页）。
- **专业排版**: 通过 python-docx 生成结构清晰的 Word 文档，支持自定义页眉、字体和页边距。

## ⚙️ 环境要求

在运行此工具前，请确保您的系统中已安装以下软件：

1.  **Python 3.11+**
2.  **python-docx 相关依赖**: 通过 `uv pip install -e .` 自动安装即可。

## 🚀 安装

1.  **克隆本仓库:**
    ```bash
    git clone https://github.com/yourusername/CodeCollate-CLI-P.git
    cd CodeCollate-CLI-P
    ```

2.  **创建并激活虚拟环境 (推荐使用 uv):**
    ```bash
    uv venv
    source .venv/bin/activate  # macOS/Linux
    # .\.venv\Scripts\activate  # Windows
    ```

3.  **安装项目及其依赖:**
    ```bash
    uv pip install -e .
    ```
    此命令将会安装项目，并使 `codecollate` 命令在您的当前环境中可用。

## 📖 使用方法

`codecollate` 命令的基本用法如下：

```bash
codecollate <SOURCE_DIR> <SOFTWARE_NAME> <VERSION> [OPTIONS]
```

### 参数说明

- `SOURCE_DIR`: **[可选]** 源代码目录路径（相对工作目录或绝对路径）。未提供时默认使用工作目录。
- `SOFTWARE_NAME`: **[必须]** 您的软件名称（如果包含空格，请使用引号）。
- `VERSION`: **[必须]** 您的软件版本号。
- `-o, --output`: **[可选]** 输出文档的目录。默认为工作目录。
- `-w, --workdir`: **[可选]** 工作目录。未提供时会交互式提示。
- `-v, --verbose`: **[可选]** 启用详细日志输出，方便调试。
- `CODECOLLATE_CONFIG_PATH`: **[环境变量]** 指向配置文件路径，优先级高于 `--config` 和工作目录默认配置。

### 示例

假设您的源代码位于工作目录 `./my-app` 下的 `src` 目录，您想为 "Super Editor v1.2" 生成文档，并输出到 `./documents` 文件夹：

```bash
codecollate src "Super Editor" "v1.2" --workdir ./my-app --output ./documents
```

命令执行成功后，您将在 `./documents` 目录下找到 `Super_Editor_v1.2_SourceCode.docx` 文件。

如果省略 `SOURCE_DIR`，工具会默认使用工作目录，并在交互式提示中补齐缺失的参数。

## 📄 许可证

本项目采用 [MIT 许可证](LICENSE)。
