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
//	               后缀为 .pdf 时先替换占位符再转换为 pdf
//	               （默认依次尝试：Word/WPS COM 自动化、pdf-renderer 子目录
//	               或 PATH 里的 office2pdf.exe、minipdf.exe、pdfitdown.exe）
//	-j / --json    JSON 字符串，key 为占位符、value 为替换后的内容（可选，
//	               不传则只做格式转换，不做替换）
//	-I / --images  需要插入的图片路径，多个用分号(;)分隔（可选）。
//	               传入时替换文档中的 {ImagesPlaceholder}，不传则删除该占位符
//	-e / --engine  PDF 转换引擎：minipdf、office2pdf、pdfitdown 或 com（可选，
//	               默认 auto，按 COM -> office2pdf -> minipdf -> pdfitdown 自动选择）
package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"image/png"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/fumiama/go-docx"
	"golang.org/x/image/bmp"
)

// main 负责解析命令行参数并调用 run 执行主流程，
// 出错时向 stderr 打印错误并以退出码 1 结束。
func main() {
	var input, output, jsonStr, images, engine string

	// flag 包不支持同一参数的长短两个名字共享配置，
	// 这里把长短两个名字绑定到同一个变量上，后出现的覆盖先出现的。
	// Go 的 flag 包同时接受 -input 和 --input 两种写法。
	flag.StringVar(&input, "input", "", "input docx file path")
	flag.StringVar(&input, "i", "", "input docx file path (shorthand)")
	flag.StringVar(&output, "output", "", "output file path (.docx or .pdf)")
	flag.StringVar(&output, "o", "", "output file path (shorthand)")
	flag.StringVar(&jsonStr, "json", "", "JSON object mapping placeholders to values")
	flag.StringVar(&jsonStr, "j", "", "JSON object (shorthand)")
	flag.StringVar(&images, "images", "", "image file paths separated by ';'")
	flag.StringVar(&images, "I", "", "image file paths (shorthand)")
	flag.StringVar(&engine, "engine", "", "pdf engine: minipdf, office2pdf, com (default auto)")
	flag.StringVar(&engine, "e", "", "pdf engine (shorthand)")
	flag.Parse()

	if err := run(input, output, jsonStr, images, engine); err != nil {
		fmt.Fprintln(os.Stderr, "error:", err)
		os.Exit(1)
	}
}

// logf 向 stderr 输出执行过程信息，不污染 stdout 的结果输出。
func logf(format string, args ...any) {
	fmt.Fprintf(os.Stderr, "[docx-converter] "+format+"\n", args...)
}

// run 是主流程：校验参数 -> 解析 JSON 与图片列表 -> 读取并替换 docx
// -> 按后缀输出 docx 或 pdf。
func run(input, output, jsonStr, images, engine string) error {
	// 必填参数校验
	if input == "" {
		return errors.New("--input is required")
	}
	if output == "" {
		return errors.New("--output is required")
	}
	logf("input: %s", input)
	logf("output: %s", output)

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
	logf("loaded %d placeholder replacement(s)", len(replacements))

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

	// 遍历文档正文的所有段落（含表格单元格内的段落），替换文本占位符。
	forEachParagraph(doc, func(p *docx.Paragraph) {
		replaceInParagraph(p, replacements)
	})

	// 处理图片占位符：传了 --images 就把 {ImagesPlaceholder}
	// 换成实际图片，没传则删除该占位符。
	if err := replaceImagePlaceholder(doc, images); err != nil {
		return err
	}

	if strings.EqualFold(filepath.Ext(output), ".pdf") {
		// 输出 pdf：先把替换后的文档写到一个临时 docx，
		// 再由 docxToPDF（office2pdf 优先，COM 回退）转成 pdf。
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

		if err := docxToPDF(tmp.Name(), output, engine); err != nil {
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

// docxToPDF 把 docx 转换为 pdf。engine 指定转换引擎：
//   - ""/"auto"：按 Word/WPS COM -> office2pdf -> minipdf -> pdfitdown
//     的顺序自动选择，前一个不可用或失败时自动尝试下一个；
//   - "minipdf" / "office2pdf" / "pdfitdown"：只用对应的命令行引擎，
//     找不到或失败直接报错，不回退；
//   - "com"：只用 Word/WPS 的 COM 自动化。
func docxToPDF(docxPath, pdfPath, engine string) error {
	start := time.Now()
	var err error
	switch strings.ToLower(engine) {
	case "", "auto":
		err = docxToPDFAuto(docxPath, pdfPath)
	case "minipdf", "office2pdf", "pdfitdown":
		err = cliEngineToPDF(strings.ToLower(engine)+".exe", docxPath, pdfPath)
	case "com":
		logf("converting to pdf with Word/WPS COM automation")
		err = docxToPDFViaWord(docxPath, pdfPath)
	default:
		return fmt.Errorf("invalid --engine %q, want minipdf, office2pdf, pdfitdown, com or auto", engine)
	}
	if err != nil {
		return err
	}
	logf("pdf conversion finished in %s", time.Since(start).Round(time.Millisecond))
	return nil
}

// docxToPDFAuto 按 Word/WPS COM -> office2pdf -> minipdf -> pdfitdown
// 的优先级自动尝试各转换引擎，全部不可用时返回错误。
func docxToPDFAuto(docxPath, pdfPath string) error {
	if wordCOMAvailable() {
		logf("converting to pdf with Word/WPS COM automation")
		if err := docxToPDFViaWord(docxPath, pdfPath); err == nil {
			return nil
		} else {
			logf("word com failed: %v; trying next engine", err)
		}
	} else {
		logf("Word/WPS COM not available, trying next engine")
	}

	var lastErr error
	for _, e := range cliEngines {
		exe := findToolExe(e.name)
		if exe == "" {
			logf("%s not found, trying next engine", e.name)
			continue
		}
		logf("converting to pdf with %s (%s)", e.name, exe)
		if err := cliToPDF(exe, e, docxPath, pdfPath); err != nil {
			logf("%s failed: %v; trying next engine", e.name, err)
			lastErr = err
			continue
		}
		return nil
	}
	if lastErr != nil {
		return lastErr
	}
	return errors.New("no pdf engine available: install Word/WPS, office2pdf, minipdf or pdfitdown")
}

// wordCOMAvailable 通过查询注册表判断 Word.Application COM 组件是否已注册
// （Microsoft Word 和 WPS Office 安装后都会注册这个 ProgID）。
// 比直接启动 COM 试转快得多，也不会在没有 Office 的机器上刷错误输出。
func wordCOMAvailable() bool {
	return exec.Command("reg", "query", `HKCR\Word.Application`).Run() == nil
}

// cliEngine 描述一个命令行 PDF 转换引擎：可执行文件名，
// 以及由输入 docx 路径和输出 pdf 路径构造命令行参数的方式
// （各工具的参数形式不同，如 office2pdf/minipdf 是位置参数，
// pdfitdown 是 -i/-o 选项）。
type cliEngine struct {
	name string
	args func(docxPath, pdfPath string) []string
}

// cliEngines 是 auto 模式下在 COM 之后依次尝试的命令行引擎。
var cliEngines = []cliEngine{
	{"office2pdf.exe", func(d, p string) []string { return []string{d, "-o", p} }},
	{"minipdf.exe", func(d, p string) []string { return []string{d, "-o", p} }},
	{"pdfitdown.exe", func(d, p string) []string { return []string{"-i", d, "-o", p} }},
}

// cliEngineToPDF 使用用户显式指定的命令行引擎转换，
// 找不到或失败时直接报错，不回退到其他引擎。
func cliEngineToPDF(name, docxPath, pdfPath string) error {
	var engine *cliEngine
	for i := range cliEngines {
		if cliEngines[i].name == name {
			engine = &cliEngines[i]
			break
		}
	}
	if engine == nil {
		return fmt.Errorf("unknown cli engine %q", name)
	}
	exe := findToolExe(name)
	if exe == "" {
		return fmt.Errorf("%s not found (put it in the pdf-renderer folder next to docx-converter.exe, or in PATH)", name)
	}
	logf("converting to pdf with %s (%s)", name, exe)
	return cliToPDF(exe, *engine, docxPath, pdfPath)
}

// findToolExe 查找名为 name 的命令行工具：先找本程序同目录下的
// pdf-renderer 子目录，再找本程序同目录，最后找 PATH。
func findToolExe(name string) string {
	if self, err := os.Executable(); err == nil {
		dir := filepath.Dir(self)
		for _, p := range []string{
			filepath.Join(dir, "pdf-renderer", name),
			filepath.Join(dir, name),
		} {
			if _, err := os.Stat(p); err == nil {
				return p
			}
		}
	}
	if p, err := exec.LookPath(name); err == nil {
		return p
	}
	return ""
}

// cliToPDF 调用命令行转换工具完成转换，命令行参数由引擎描述构造。
func cliToPDF(exe string, engine cliEngine, docxPath, pdfPath string) error {
	docxAbs, err := filepath.Abs(docxPath)
	if err != nil {
		return err
	}
	pdfAbs, err := filepath.Abs(pdfPath)
	if err != nil {
		return err
	}

	// 先删除旧的输出文件，避免把上次运行的残留误认为本次成功
	os.Remove(pdfAbs)

	cmd := exec.Command(exe, engine.args(docxAbs, pdfAbs)...)
	out, cmdErr := cmd.CombinedOutput()
	if _, err := os.Stat(pdfAbs); err != nil {
		name := filepath.Base(exe)
		if cmdErr != nil {
			return fmt.Errorf("%s: %w: %s", name, cmdErr, out)
		}
		return fmt.Errorf("%s: pdf not created: %s", name, out)
	}
	return nil
}

// docxToPDFViaWord 通过 COM 自动化把 docx 转换为 pdf。
// 原理是启动一个隐藏的 Word 进程，打开文档后另存为 PDF（wdFormatPDF = 17），
// 最后退出 Word。依赖本机安装的 Word.Application COM 组件
// （Microsoft Word 或 WPS Office 均可），仅适用于 Windows。
func docxToPDFViaWord(docxPath, pdfPath string) error {
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
$word.DisplayAlerts = 0
try {
    # Open 的后几个参数：ConfirmConversions=$false, ReadOnly=$true, AddToRecentFiles=$false
    # 只读打开既省去最近文档簿记，也避免 Word 锁定源文件
    $doc = $word.Documents.Open('%s', $false, $true, $false)
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

// imagePlaceholders 是可被 --images 替换的图片占位符。
// 单复数两种写法都认，防止模板里写错。
var imagePlaceholders = []string{"{ImagesPlaceholder}", "{ImagePlaceholder}"}

// forEachParagraph 遍历文档正文的所有段落，包括表格单元格里的段落。
// 正文顶层元素只有两类：段落（Paragraph）和表格（Table）；
// 表格需要逐行、逐单元格找到其中的段落。
func forEachParagraph(doc *docx.Docx, fn func(p *docx.Paragraph)) {
	for _, item := range doc.Document.Body.Items {
		switch v := item.(type) {
		case *docx.Paragraph:
			fn(v)
		case *docx.Table:
			for _, row := range v.TableRows {
				for _, cell := range row.TableCells {
					for _, p := range cell.Paragraphs {
						fn(p)
					}
				}
			}
		}
	}
}

// replaceImagePlaceholder 处理文档中的图片占位符：
//   - images 非空（形如 "a.png;b.bmp"）：先把占位符文本删掉，
//     再按顺序把每张图作为 inline drawing 插入到占位符所在段落；
//   - images 为空：只删除占位符文本。
func replaceImagePlaceholder(doc *docx.Docx, images string) error {
	var paths []string
	for _, p := range strings.Split(images, ";") {
		if p = strings.TrimSpace(p); p != "" {
			paths = append(paths, p)
		}
	}
	if len(paths) > 0 {
		logf("loading %d image(s): %s", len(paths), strings.Join(paths, ", "))
	} else {
		logf("no --images given, image placeholders will be removed")
	}

	// 删除占位符用的替换表：占位符 -> 空串
	remove := make(map[string]string, len(imagePlaceholders))
	for _, ph := range imagePlaceholders {
		remove[ph] = ""
	}

	var (
		firstErr error
		datas    [][]byte
		loaded   bool
	)
	forEachParagraph(doc, func(p *docx.Paragraph) {
		// 占位符可能被拆到多个 run 里，拼接整段文字再判断是否存在，
		// 避免漏掉跨节点拆分的情况。
		var joined strings.Builder
		for _, t := range textNodes(p) {
			joined.WriteString(t.Text)
		}
		hasPlaceholder := false
		for _, ph := range imagePlaceholders {
			if strings.Contains(joined.String(), ph) {
				hasPlaceholder = true
				break
			}
		}

		// 删除占位符文本（replaceInParagraph 已处理跨 run 拆分）
		replaceInParagraph(p, remove)

		if len(paths) == 0 || !hasPlaceholder {
			return
		}

		// 图片只在第一次遇到占位符时加载（文档里可能没有占位符，
		// 避免白做解码）；多张图并行读文件/转码，大图能省一半时间。
		if !loaded {
			loaded = true
			datas = make([][]byte, len(paths))
			errs := make([]error, len(paths))
			var wg sync.WaitGroup
			for i, path := range paths {
				wg.Add(1)
				go func(i int, path string) {
					defer wg.Done()
					datas[i], errs[i] = loadImageData(path)
				}(i, path)
			}
			wg.Wait()
			for _, err := range errs {
				if err != nil {
					firstErr = err
					break
				}
			}
		}
		if firstErr != nil {
			return
		}

		for i, data := range datas {
			if _, err := p.AddInlineDrawing(data); err != nil {
				firstErr = fmt.Errorf("add image %s: %w", paths[i], err)
				return
			}
		}
		logf("inserted %d image(s) at image placeholder", len(datas))
	})
	return firstErr
}

// loadImageData 读取图片文件并返回可直接写入 docx 的字节。
// go-docx 依赖 imgsz 探测图片尺寸，而 imgsz 只支持 png/jpg/gif/webp，
// 不支持 bmp，所以 bmp 文件先在内存里转成 png 再交给 go-docx。
func loadImageData(path string) ([]byte, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("read image %s: %w", path, err)
	}
	if !strings.EqualFold(filepath.Ext(path), ".bmp") {
		return data, nil
	}
	logf("converting bmp to png: %s", path)
	img, err := bmp.Decode(bytes.NewReader(data))
	if err != nil {
		return nil, fmt.Errorf("decode bmp %s: %w", path, err)
	}
	var buf bytes.Buffer
	// 用 BestSpeed 而不是默认压缩级别：2048x2048 的图编码时间从 ~1s 降到 ~0.2s，
	// 体积只从 1.9MB 涨到 2.6MB；这些字节最终嵌进 docx，Word 转 pdf 时还会重压缩。
	enc := png.Encoder{CompressionLevel: png.BestSpeed}
	if err := enc.Encode(&buf, img); err != nil {
		return nil, fmt.Errorf("convert bmp %s to png: %w", path, err)
	}
	return buf.Bytes(), nil
}
