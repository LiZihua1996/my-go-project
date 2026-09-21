package main

import (
	"flag"
	"fmt"
	"os"
	"todo"
)

const todo_file_name = "todo.json"

func main() {
	var add = flag.String("add", "", "Task to be included in the TodoList")
	var list = flag.Bool("list", false, "List all tasks in the TodoList")
	var complete = flag.Int("complete", -1, "Item to be complete")
	//var clear = flag.Bool("clear", false, "Clear completed tasks");

	flag.Parse()

	todo_list := &todo.TodoList{}
	if err := todo_list.Load(todo_file_name); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}

	switch {
	case *list:
		// 列出所有待办事项
		fmt.Print(todo_list)
	case *complete >= 0:
		// 完成待办事项
		if err := todo_list.Complete(*complete); err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}

		// 保存待办事项列表到文件
		if err := todo_list.Save(todo_file_name); err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}

		// 列出所有待办事项
		fmt.Print(todo_list)

	case *add != "":
		// 添加待办事项
		todo_list.Add(*add)

		// 保存待办事项列表到文件
		if err := todo_list.Save(todo_file_name); err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}

		// 列出所有待办事项
		fmt.Print(todo_list)

	default:
		// 无有效参数
		fmt.Println("Invalid option")
		os.Exit(1)
	}
}
