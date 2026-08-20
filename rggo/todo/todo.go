package todo

import (
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"time"
)

// todo_item 待办事项项
type todo_item struct {
	Description  string
	Done         bool
	CreatedAt    time.Time
	CompleteedAt time.Time
}

// TodoList 待办事项列表
type TodoList []todo_item

// Add 添加待办事项
func (l *TodoList) Add(task string) {
	var t = todo_item{
		Description:  task,
		Done:         false,
		CreatedAt:    time.Now(),
		CompleteedAt: time.Time{},
	}
	*l = append(*l, t)
}

// Complete 完成待办事项
func (l *TodoList) Complete(i int) error {
	var list = *l
	if i < 0 || i >= len(list) {
		return fmt.Errorf("index out of range")
	}
	list[i].Done = true
	list[i].CompleteedAt = time.Now()
	return nil
}

// Delete 删除待办事项
func (l *TodoList) Delete(i int) error {
	var list = *l
	if i < 0 || i >= len(list) {
		return fmt.Errorf("index out of range")
	}
	*l = append(list[:i], list[i+1:]...)
	return nil
}

// Save 保存待办事项列表到文件
func (l *TodoList) Save(filname string) error {
	// 将 List（待办事项列表）序列化为 JSON 格式的字节数据
	var js, err = json.Marshal(*l)
	if err != nil {
		return err
	}
	return os.WriteFile(filname, js, 0644)
}

// Load 从文件加载待办事项列表
func (l *TodoList) Load(filname string) error {
	// 从文件中读取 JSON 格式的字节数据
	var bytes, err = os.ReadFile(filname)
	if err != nil {
		if errors.Is(err, os.ErrNotExist) {
			return nil
		}
		return err
	}

	if len(bytes) == 0 {
		return nil
	}

	return json.Unmarshal(bytes, l)
}
