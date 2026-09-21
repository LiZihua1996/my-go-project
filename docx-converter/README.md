# docx-converter

一个命令行工具：读取 docx 模板文件，替换其中的占位符（如 `{Name}`、`{Date}`），并按输出后缀生成 docx 或 pdf 文件。

## 功能

- 基于 [go-docx](https://github.com/fumiama/go-docx) 解析 docx，替换正文段落和表格中的占位符
- 占位符被 Word 拆分到多个 run 时也能正确替换
- 输出 `.docx`：直接生成替换后的文档
- 输出 `.pdf`：优先使用同目录或 PATH 中的 [office2pdf](https://github.com/developer0hye/office2pdf)
  （独立 exe，不依赖 Office/WPS，约 2s）；找不到时回退到 Word/WPS 的 COM 自动化
  （仅 Windows，需安装 Microsoft Word 或 WPS Office）

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
| `--images` | `-I` | 否 | 需要插入的图像的路径用分号(;)分隔

## 使用示例

PowerShell（建议用单引号包裹 JSON):

```powershell
.\docx-converter.exe -i template.docx -o output.docx -j '{"{Name}":"Mike","{Date}":"2026-08-01"}'
.\docx-converter.exe --input template.docx --output output.pdf --json '{"{Name}":"Mike"}'
```

说明：

- JSON 必须是合法格式（key 与 value 之间是冒号），单引号写法会被自动兼容
- 图片插入：在模板中放置 `{ImagesPlaceholder}`（兼容 `{ImagePlaceholder}` 写法），
  传入 `--images` 时占位符按顺序替换为所给图片，不传时占位符被删除；
  bmp 图片会先转换为 png 再插入（go-docx 依赖的 imgsz 不支持 bmp）
- PDF 转换引擎：把 `office2pdf.exe` 放到 docx-converter.exe 同目录（推荐，
  可从 <https://github.com/developer0hye/office2pdf/releases> 下载），
  即可享受无 Office 依赖的快速转换；没有它则自动回退到本机 Word/WPS 的 COM 自动化
