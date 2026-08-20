package main_test

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
)

var bin_name = "todo"
var filename = "todo.json"

func TestMain(m *testing.M) {
	// 测试代码
	fmt.Println("Building tool...")
	if runtime.GOOS == "windows" {
		bin_name += ".exe"
	}

	// 执行构建命令
	var build = exec.Command("go", "build", "-o", bin_name)
	if err := build.Run(); err != nil {
		fmt.Fprintf(os.Stderr, "Failed to build tool: %s: %s", bin_name, err)
		os.Exit(1)
	}

	fmt.Println("Running tests...")
	var result = m.Run() //执行此测试文件中的所有测试函数

	fmt.Println("Cleaning up...")
	os.Remove(bin_name)
	os.Remove(filename)
	os.Exit(result)
}

func TestTodoCLI(t *testing.T) {
	var task = "test task"

	var dir, err = os.Getwd() // 获取当前工作目录
	if err != nil {
		t.Fatal(err)
	}

	var cmd_path = filepath.Join(dir, bin_name) // 构建命令路径

	//调用todo.exe添加新任务
	t.Run("AddNewTask", func(t *testing.T) {
		var cmd = exec.Command(cmd_path, strings.Split(task, " ")...) // 构建命令
		if err := cmd.Run(); err != nil {
			t.Fatal(err)
		}
	})

	//调用todo.exe看看输出的任务名称和刚才添加的任务名称是否一致
	t.Run("ListTasks", func(t *testing.T) {
		var cmd = exec.Command(cmd_path)
		var out, err = cmd.CombinedOutput()
		if err != nil {
			t.Fatal(err)
		}

		var expected = task + "\n"
		if string(out) != expected {
			t.Errorf("Expected %q, got %q instead", expected, string(out))
		}
	})
}
