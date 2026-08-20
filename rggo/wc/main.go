package main

import (
	"bufio"
	"flag"
	"fmt"
	"io"
	"os"
)

// count 统计从 r 读取的文本中的单词数量
// 它使用 bufio.Scanner 以单词为单位进行扫描
func count(r io.Reader, count_lines bool) (int, error) {
	// 创建一个新的扫描器来读取输入
	var scanner = bufio.NewScanner(r)
	if !count_lines {
		// 设置扫描器的分割函数为按单词分割
		scanner.Split(bufio.ScanWords)
	} else {
		// 设置扫描器的分割函数为按行分割
		scanner.Split(bufio.ScanLines) //ScanLines是默认的分割函数
	}
	// 初始化单词计数器
	var count = 0

	// 逐个扫描单词，每扫描到一个单词计数器加一
	for scanner.Scan() {
		count++
	}
	// 检查扫描过程中是否发生了错误（EOF除外）
	if err := scanner.Err(); err != nil {
		return 0, err
	}
	return count, nil
}

func main() {
	// 解析命令行标志
	// flag.Bool 的三个参数依次为：标志名 "l"、默认值 false、帮助信息 "Count lines"
	var lines = flag.Bool("l", false, "Count lines")
	flag.Parse()
	// 从标准输入读取数据，统计单词数量并输出
	count, err := count(os.Stdin, *lines)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	fmt.Println(count)
}
