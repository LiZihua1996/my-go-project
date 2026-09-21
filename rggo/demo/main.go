package main

import (
	"demo/book"
	"fmt"
)

func main() {
	fmt.Println("Hello, World!")
	var book = book.Book{
		Title:  "Go in Action",
		Author: "Chen Yue",
	}
	fmt.Println(book)
}
