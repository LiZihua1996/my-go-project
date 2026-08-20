// todo_test 包对 todo 包的功能进行单元测试
package todo_test

import (
	"os"
	"testing"
	"todo"
)

// TestAdd 测试 Add 方法：添加一条任务后，验证其 Description 字段是否正确
func TestAdd(t *testing.T) {
	l := todo.TodoList{}
	task_name := "New Task"
	l.Add(task_name)

	// 验证添加后的任务描述与传入值一致
	if l[0].Description != task_name {
		t.Errorf("Add failed, expected %s, got %s", task_name, l[0].Description)
	}
}

// TestComplete 测试 Complete 方法：先添加任务，再标记完成，验证 Done 状态切换
func TestComplete(t *testing.T) {
	l := todo.TodoList{}

	task_name := "New Task"
	l.Add(task_name)

	// 确认添加成功
	if l[0].Description != task_name {
		t.Errorf("Add failed, expected %s, got %s", task_name, l[0].Description)
	}

	// 刚添加的任务应该是未完成状态
	if true == l[0].Done {
		t.Errorf("Complete failed, expected false, got true")
	}

	// 标记第 0 项为已完成
	l.Complete(0)

	// 验证 Done 已变为 true
	if false == l[0].Done {
		t.Errorf("Complete failed, expected true, got false")
	}
}

// TestDelete 测试 Delete 方法：添加 3 条任务，删除索引 1，验证长度和剩余元素正确性
func TestDelete(t *testing.T) {
	l := todo.TodoList{}

	task_descriptions := []string{"New Task", "New Task 2", "New Task 3"}

	// 批量添加任务
	for _, task_description := range task_descriptions {
		l.Add(task_description)
	}

	// 确认添加成功
	if l[0].Description != task_descriptions[0] {
		t.Errorf("Delete failed, expected %s, got %s", task_descriptions[0], l[0].Description)
	}

	// 删除索引 1（即 "New Task 2"）
	l.Delete(1)

	// 删除后列表长度应为 2
	if len(l) != 2 {
		t.Errorf("Delete failed, expected 2, got %d", len(l))
	}

	// 删除后索引 1 应为原索引 2 的 "New Task 3"
	if l[1].Description != task_descriptions[2] {
		t.Errorf("Delete failed, expected %s, got %s", task_descriptions[2], l[1].Description)
	}
}

// testSaveAndLoad 测试 Save 和 Load 方法：将列表保存到临时文件后重新加载，验证数据一致性
// 注意：函数名小写开头（testSaveAndLoad），不会被 go test 自动执行，需通过 TestXxx 调用
func testSaveAndLoad(t *testing.T) {
	l1 := todo.TodoList{} // 用于保存的列表
	l2 := todo.TodoList{} // 用于加载的列表

	task_description := "New Task"
	l1.Add(task_description)

	// 确认添加成功
	if l1[0].Description != task_description {
		t.Errorf("Add failed, expected %s, got %s", task_description, l1[0].Description)
	}

	// 在系统默认临时目录创建临时文件
	tf, err := os.CreateTemp("", "")

	if err != nil {
		t.Errorf("Error create temp file, err: %s", err)
	}

	// 测试结束后删除临时文件
	defer os.Remove(tf.Name())

	// 将 l1 保存到临时文件
	if err := l1.Save(tf.Name()); err != nil {
		t.Fatalf("Error save todo list to file, err: %s", err)
	}

	// 从临时文件加载到 l2
	if err := l2.Load(tf.Name()); err != nil {
		t.Fatalf("Error load todo list from file, err: %s", err)
	}

	// 验证加载后的数据与原数据一致
	if l2[0].Description != task_description {
		t.Errorf("Load failed, expected %s, got %s", task_description, l2[0].Description)
	}

}
