# docx-converter

一个命令行工具：读取 docx 模板文件，替换其中的占位符（如 `{Name}`、`{Date}`），并按输出后缀生成 docx 或 pdf 文件。

## 功能

- 基于 [wordZero](https://github.com/zerx-lab/wordZero) 解析 docx，替换正文段落和表格中的占位符
  （注意：该仓库的 Go module 路径是 `github.com/ZeroHawkeye/wordZero`）
- 占位符被 Word 拆分到多个 run 时也能正确替换
- 输出 `.docx`：直接生成替换后的文档
- 输出 `.pdf`：可用 `--engine` 指定转换引擎（`minipdf` / `office2pdf` / `pdfitdown` / `com`），
  默认 `auto` 时依次尝试以下引擎，前一个不可用时自动回退到下一个：
  1. Word/WPS 的 COM 自动化（仅 Windows，需安装 Microsoft Word 或 WPS Office，排版还原度最高）
  2. pdf-renderer 子目录或 PATH 中的 [office2pdf](https://github.com/developer0hye/office2pdf)（独立 exe，不依赖 Office/WPS，约 2s）
  3. pdf-renderer 子目录或 PATH 中的 [MiniPdf](https://github.com/mini-software/MiniPdf) Rust CLI
     （`cargo install minipdf-cli`，不依赖 Office/WPS，支持图片和基础表格）
  4. pdf-renderer 子目录或 PATH 中的 [PdfItDown](https://github.com/AstraBert/PdfItDown) CLI
     （不依赖 Office/WPS，docx 转换基于 office2pdf crate，还支持 md/html/图片等格式）

## 构建

```sh
go build -o docx-converter.exe .
```

## 命令行参数

| 参数 | 简写 | 必填 | 说明 |
| --- | --- | --- | --- |
| `--input` | `-i` | 是 | 输入的 docx 文件路径 |
| `--output` | `-o` | 是 | 输出文件路径，后缀 `.docx` 生成 docx，`.pdf` 转换为 pdf |
| `--json` | `-j` | 否 | JSON 字符串，key 为占位符、value 为替换内容；不传则只做格式转换 |
| `--images` | `-I` | 否 | 需要插入的图像的路径用分号(;)分隔 |
| `--engine` | `-e` | 否 | PDF 转换引擎：`minipdf`、`office2pdf`、`pdfitdown`、`com`；默认 `auto`，按 COM -> office2pdf -> minipdf -> pdfitdown 自动选择。显式指定时找不到或失败会直接报错，不回退 |

## 使用示例

PowerShell(建议用单引号包裹 JSON):

```powershell
.\docx-converter.exe -i template.docx -o output.docx -j '{"{Name}":"Mike","{Date}":"2026-08-01"}'
.\docx-converter.exe --input template.docx --output output.pdf --json '{"{Name}":"Mike"}'
```

说明：

- JSON 必须是合法格式（key 与 value 之间是冒号），单引号写法会被自动兼容
- 图片插入：在模板中放置 `{ImagesPlaceholder}`（兼容 `{ImagePlaceholder}` 写法），
  传入 `--images` 时占位符按顺序替换为所给图片，不传时占位符被删除；
  插入的图片宽度按页面可用宽度（页宽减左右边距）的 1/3 缩放、高度等比换算，
  每行最多排 3 张，多出来的自动另起一行；
  bmp 图片会先转换为 png 再插入（wordZero 只支持 png/jpg/gif，不支持 bmp）
- PDF 转换引擎：装了 Microsoft Word 或 WPS Office 时默认优先用 COM 自动化；
  没有 Office 环境时，把 `office2pdf.exe` 放到 docx-converter.exe 旁的 pdf-renderer 文件夹
  （可从 <https://github.com/developer0hye/office2pdf/releases> 下载），
  或放入 `minipdf.exe`（用 `cargo install minipdf-cli` 安装后从 `~/.cargo/bin` 拷贝）、
  `pdfitdown.exe`（可从 <https://github.com/AstraBert/PdfItDown> 获取），
  即可享受无 Office 依赖的转换
