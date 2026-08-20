package main

import (
	"fmt"
	"os"
	"strings"
	"todo"
)

const todo_file_name = "todo.json"

func main() {
	todo_list := todo.TodoList{}
	if err := todo_list.Load(todo_file_name); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}

	switch {
	case len(os.Args) == 1:
		for _, task := range todo_list {
			fmt.Printf("%s\n", task.Description)
		}
	default:
		var item = strings.Join(os.Args[1:], " ")
		todo_list.Add(item)

		if err := todo_list.Save(todo_file_name); err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}
	}
}
