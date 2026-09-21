// docx-converter 是一个命令行工具，用于读取 docx 模板文件，
// 将其中的占位符（如 {Name}、{Date}）替换为指定内容，
// 然后按输出文件后缀生成 docx 或 pdf 文件。
//
// 用法示例（PowerShell 中建议用单引号包裹 JSON）：
//
//	docx-converter.exe -i template.docx -o output.docx -j '{"{Name}":"Mike","{Date}":"2026-08-01"}'
//	docx-converter.exe -i template.docx -o output.pdf  -j '{"{Name}":"Mike"}'
//
// 参数说明：
//
//	-i / --input   输入的 docx 文件路径（必填）
//	-o / --output  输出文件路径（必填），后缀为 .docx 时直接生成 docx，
//	               后缀为 .pdf 时先替换占位符再通过 Word 转换为 pdf
//	-j / --json    JSON 字符串，key 为占位符、value 为替换后的内容（可选，
//	               不传则只做格式转换，不做替换）
package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"strings"

	"github.com/fumiama/go-docx"
)

// main 负责解析命令行参数并调用 run 执行主流程，
// 出错时向 stderr 打印错误并以退出码 1 结束。
func main() {
	var input, output, jsonStr string

	// flag 包不支持同一参数的长短两个名字共享配置，
	// 这里把长短两个名字绑定到同一个变量上，后出现的覆盖先出现的。
	// Go 的 flag 包同时接受 -input 和 --input 两种写法。
	flag.StringVar(&input, "input", "", "input docx file path")
	flag.StringVar(&input, "i", "", "input docx file path (shorthand)")
	flag.StringVar(&output, "output", "", "output file path (.docx or .pdf)")
	flag.StringVar(&output, "o", "", "output file path (shorthand)")
	flag.StringVar(&jsonStr, "json", "", "JSON object mapping placeholders to values")
	flag.StringVar(&jsonStr, "j", "", "JSON object (shorthand)")
	flag.Parse()

	if err := run(input, output, jsonStr); err != nil {
		fmt.Fprintln(os.Stderr, "error:", err)
		os.Exit(1)
	}
}

// run 是主流程：校验参数 -> 解析 JSON -> 读取并替换 docx -> 按后缀输出 docx 或 pdf。
func run(input, output, jsonStr string) error {
	// 必填参数校验
	if input == "" {
		return errors.New("--input is required")
	}
	if output == "" {
		return errors.New("--output is required")
	}

	// 解析 JSON 替换表，例如 {"{Name}":"Mike","{Date}":"2026-08-01"}
	replacements := make(map[string]string)
	if jsonStr != "" {
		if err := json.Unmarshal([]byte(jsonStr), &replacements); err != nil {
			// Windows 命令行里双引号转义麻烦，用户常会写成单引号 JSON，
			// 例如 {'{Name}':'Mike'}。严格解析失败时做一次单引号替换再试。
			// 注意：此兜底不支持 value 中本身包含单引号的情况。
			tolerant := strings.ReplaceAll(jsonStr, "'", "\"")
			if err2 := json.Unmarshal([]byte(tolerant), &replacements); err2 != nil {
				// 两种写法都失败时报错，并回显收到的原始字符串，
				// 方便排查 shell 引号转义问题。
				return fmt.Errorf("invalid --json: %w (received: %s)", err, jsonStr)
			}
		}
	}

	// 根据输出文件后缀决定输出格式，只支持 .docx 和 .pdf
	switch ext := strings.ToLower(filepath.Ext(output)); ext {
	case ".docx", ".pdf":
	default:
		return fmt.Errorf("unsupported output extension %q, want .docx or .pdf", ext)
	}

	// 读取输入 docx（预先去掉 zip 中的目录条目，见 readDocxSkipDirEntries）
	data, err := readDocxSkipDirEntries(input)
	if err != nil {
		return err
	}

	// 用 go-docx 解析文档。Parse 需要 io.ReaderAt 和文件大小，
	// 这里用内存中的字节切片构造。
	doc, err := docx.Parse(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return err
	}

	// 遍历文档正文的顶层元素，逐个替换占位符。
	// 顶层元素只有两类：段落（Paragraph）和表格（Table）；
	// 表格需要逐行、逐单元格找到其中的段落再替换。
	for _, item := range doc.Document.Body.Items {
		switch v := item.(type) {
		case *docx.Paragraph:
			replaceInParagraph(v, replacements)
		case *docx.Table:
			for _, row := range v.TableRows {
				for _, cell := range row.TableCells {
					for _, p := range cell.Paragraphs {
						replaceInParagraph(p, replacements)
					}
				}
			}
		}
	}

	if strings.EqualFold(filepath.Ext(output), ".pdf") {
		// 输出 pdf：先把替换后的文档写到一个临时 docx，
		// 再通过 Word COM 把临时 docx 另存为 pdf。
		tmp, err := os.CreateTemp("", "docx-converter-*.docx")
		if err != nil {
			return err
		}
		// 无论成败，函数返回时删除临时文件
		defer os.Remove(tmp.Name())

		if _, err := doc.WriteTo(tmp); err != nil {
			tmp.Close()
			return err
		}
		// 必须先关闭文件句柄，否则 Windows 上 Word 可能打不开它
		if err := tmp.Close(); err != nil {
			return err
		}

		if err := docxToPDF(tmp.Name(), output); err != nil {
			return err
		}
	} else {
		// 输出 docx：直接把替换后的文档写到目标路径
		out, err := os.Create(output)
		if err != nil {
			return err
		}
		if _, err := doc.WriteTo(out); err != nil {
			out.Close()
			return err
		}
		if err := out.Close(); err != nil {
			return err
		}
	}

	fmt.Println("saved to", output)
	return nil
}

// docxToPDF 通过 Microsoft Word 的 COM 自动化接口把 docx 转换为 pdf。
// 原理是启动一个隐藏的 Word 进程，打开文档后另存为 PDF（wdFormatPDF = 17），
// 最后退出 Word。依赖本机安装 Microsoft Word，仅适用于 Windows。
func docxToPDF(docxPath, pdfPath string) error {
	// COM 接口要求使用绝对路径
	docxAbs, err := filepath.Abs(docxPath)
	if err != nil {
		return err
	}
	pdfAbs, err := filepath.Abs(pdfPath)
	if err != nil {
		return err
	}

	// 注意：Word 有时会在 Quit 时崩溃（HRESULT 0x800706BE），
	// 此时 PDF 其实已经生成完毕，但崩溃留下的错误记录会让
	// powershell -Command 以退出码 1 结束。因此：
	//  1. Quit 用 try/catch 包住，忽略崩溃；
	//  2. 转换成功与否不看退出码，而是检查 PDF 文件是否真的生成了。
	script := fmt.Sprintf(`
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open('%s')
    $doc.SaveAs([ref]'%s', [ref]17)
    $doc.Close()
} finally {
    try { $word.Quit() } catch {}
}
`, psEscape(docxAbs), psEscape(pdfAbs))

	// 先删除旧的输出文件，避免把上次运行的残留误认为本次成功
	os.Remove(pdfAbs)

	// -NoProfile：不加载用户配置，启动更快、行为更稳定
	// -NonInteractive：不弹交互提示，遇到错误直接失败
	cmd := exec.Command("powershell", "-NoProfile", "-NonInteractive", "-Command", script)
	out, cmdErr := cmd.CombinedOutput()
	if _, err := os.Stat(pdfAbs); err != nil {
		// PDF 没生成才是真正的失败，此时带上 powershell 的输出方便排查
		if cmdErr != nil {
			return fmt.Errorf("word com: %w: %s", cmdErr, out)
		}
		return fmt.Errorf("word com: pdf not created: %s", out)
	}
	return nil
}

// psEscape 转义 PowerShell 单引号字符串中的单引号：
// 在 PowerShell 里，单引号字符串内的单引号要写成两个单引号。
func psEscape(s string) string {
	return strings.ReplaceAll(s, "'", "''")
}

// readDocxSkipDirEntries 读取 docx 文件，并把其中的 zip 目录条目
// （名字以 / 结尾的条目，如 "_rels/"、"word/"）剔除后返回新的字节流。
//
// docx 本质是一个 zip 压缩包。很多工具生成的 docx 会包含目录条目，
// 而当前版本的 go-docx 会把这些目录条目记入文件列表，回写时尝试
// 打开它们并报错 "open _rels/: invalid argument"。
// 这里预先重写一遍 zip，跳过目录条目，规避这个问题。
func readDocxSkipDirEntries(path string) ([]byte, error) {
	r, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer r.Close()

	var buf bytes.Buffer
	w := zip.NewWriter(&buf)
	for _, f := range r.File {
		// 跳过目录条目
		if strings.HasSuffix(f.Name, "/") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			return nil, err
		}
		wc, err := w.Create(f.Name)
		if err != nil {
			rc.Close()
			return nil, err
		}
		_, err = io.Copy(wc, rc)
		rc.Close()
		if err != nil {
			return nil, err
		}
	}
	if err := w.Close(); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// replaceInParagraph 在一个段落内替换所有占位符。
//
// docx 的段落由若干 run（*docx.Run）组成，run 内的 *docx.Text 才是文字。
// Word 经常会把同一句话拆到多个 run 里（例如改过一次格式），
// 所以占位符可能横跨多个文本节点，例如 "{Na" + "me}"。
//
// 替换分两步：
//  1. 先在每个文本节点内部单独替换（覆盖最常见的不拆分情况，
//     这样不会改变任何格式）；
//  2. 如果整段拼起来仍然含有占位符，说明它被拆到了多个 run 里，
//     这时把整段替换后的文字塞进第一个文本节点、清空其余节点。
//     兜底方案会丢失段落中间 run 的独立格式，但能保证占位符被替换掉。
func replaceInParagraph(p *docx.Paragraph, replacements map[string]string) {
	texts := textNodes(p)

	// 第一步：逐节点替换
	for _, t := range texts {
		for old, new := range replacements {
			t.Text = strings.ReplaceAll(t.Text, old, new)
		}
	}

	// 拼接整段文字，检查是否还有跨节点的占位符
	var joined strings.Builder
	for _, t := range texts {
		joined.WriteString(t.Text)
	}
	full := joined.String()

	var dirty bool
	for old, new := range replacements {
		if strings.Contains(full, old) {
			full = strings.ReplaceAll(full, old, new)
			dirty = true
		}
	}

	// 第二步：存在跨节点占位符时，整段合并到第一个文本节点
	if dirty && len(texts) > 0 {
		texts[0].Text = full
		for _, t := range texts[1:] {
			t.Text = ""
		}
	}
}

// textNodes 返回段落中所有 run 里的文本节点（*docx.Text），
// 按文档顺序排列。
func textNodes(p *docx.Paragraph) []*docx.Text {
	var texts []*docx.Text
	for _, child := range p.Children {
		if run, ok := child.(*docx.Run); ok {
			for _, rc := range run.Children {
				if t, ok := rc.(*docx.Text); ok {
					texts = append(texts, t)
				}
			}
		}
	}
	return texts
}
