# PDF 2 TXT / DOCX 转换工具

将 PDF 试卷 / 文档一键转换为 TXT 和带 Word 原生公式（OMML）的 DOCX。

识别步骤：把 PDF 每一页渲染成高清图片 → 调用本地视觉大模型（也可切换云端 API）逐页识别文字与公式 → 忠实地按题目顺序还原题干与选项 → 生成 `.txt` 与 `.docx`。

## 特性

- <img src="assets/icon-check.svg" width="16" height="16" align="center" alt=""> 一键转换：拖拽 PDF 到 bat 或命令行运行
- <img src="assets/icon-check.svg" width="16" height="16" align="center" alt=""> 公式转成 **Word 原生公式对象**，可在 Word 中直接点击编辑
- <img src="assets/icon-check.svg" width="16" height="16" align="center" alt=""> 支持本地模型（Ollama）与 OpenAI 兼容云端 API 双配置
- <img src="assets/icon-check.svg" width="16" height="16" align="center" alt=""> 忠实还原题目顺序、题干、选项

## 项目结构

```
pdf2txt-docx/
├── cnvpdf.py          # 主脚本
├── config.json        # 配置（多 profile：本地 / 云端）
├── 一键转换.bat       # Windows 一键入口（拖拽 PDF 到该文件）
├── requirements.txt   # Python 依赖
└── README.md
```

## 环境要求

- Python 3.9+
- 本机装有 Microsoft Office（用于 `MML2OMML.XSL` 把 MathML 转成原生公式）
- 使用本地模型时需安装 [Ollama](https://ollama.com)

### 安装依赖

```bash
pip install -r requirements.txt
```

### 准备本地视觉模型（推荐，免费、离线）

本工具默认使用本地视觉模型 `qwen3-vl:4b`：

```bash
ollama pull qwen3-vl:4b
```

> 模型体积约 4GB，需联网下载一次。也可替换为其他支持图片的模型（在 `config.json` 修改 `model`）。

## 使用方法

### 方式一：拖拽（Windows）

直接把 PDF 文件**拖到 `一键转换.bat` 图标上**松开即可，自动生成同名 `.txt` 和 `.docx`。

> 注：bat 里写的是本机 Python 路径，若在其他电脑使用，请按需修改为已安装依赖的 python 路径。

### 方式二：命令行

```bash
# 默认输出到 PDF 同目录
python cnvpdf.py "试卷.pdf"

# 指定输出目录 / 切换配置档
python cnvpdf.py "试卷.pdf" --out "D:\输出"
python cnvpdf.py "试卷.pdf" --profile deepseek
```

## 配置说明（config.json）

```json
{
  "active": "local",
  "profiles": {
    "local": {
      "base_url": "http://localhost:11434/v1",
      "model": "qwen3-vl:4b",
      "api_key": "ollama",
      "native": "ollama",
      "num_ctx": 32768,
      "timeout": 600
    },
    "deepseek": {
      "base_url": "https://api.deepseek.com/v1",
      "model": "deepseek-v4-pro",
      "api_key": "sk-在这里填你的key",
      "timeout": 600
    }
  },
  "timeout": 600
}
```

- `active`：当前启用的配置档
- `local`：本地 Ollama 视觉模型，无需联网、免费
- `deepseek`：云端 API 示例，请填入你自己的 key
- `native`: `"ollama"` 时走 Ollama 原生接口，可调大 `num_ctx` 上下文

## 输出说明

- `.txt`：纯文本，含全部题干与选项（公式以 LaTeX 形式呈现，便于阅读）
- `.docx`：按题目顺序排版，公式为 Word 原生公式对象，可直接编辑

## 常见问题

- **报 `ModuleNotFoundError: No module named 'pymupdf'`**：未安装依赖，执行 `pip install -r requirements.txt`
- **公式识别不准**：`qwen3-vl:4b` 是 4B 小模型，复杂公式可能识别有限，可换更强的视觉模型
- **生成的文件在哪**：默认在 PDF 同一目录，文件名相同

## License

MIT
