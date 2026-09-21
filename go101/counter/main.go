package main

import (
	"bufio"
	"fmt"
	"io"
	"log"
	"os"
)

func main() {
	var filename = "words.txt"

	var file, err = os.Open(filename)
	if err != nil {
		log.Fatalln("Failed to read file: ", err)
	}

	word_count := CountWords(file)
	fmt.Println(word_count)

	os.Exit(0)
}

func CountWords(file io.Reader) int {
	word_count := 0

	scanner := bufio.NewScanner(file)
	scanner.Split(bufio.ScanWords)

	for scanner.Scan() {
		word_count++
	}

	if err := scanner.Err(); err != nil {
		log.Fatalln("Scanner error:", err)
	}
	return word_count
}

// func CountWordsInFile(file *os.File) int {
// 	word_count := 0
// 	is_inside_word := false

// 	const BUFFER_SIZE = 4096
// 	var buffer = make([]byte, BUFFER_SIZE)
// 	var leftover = []byte{}

// 	for {
// 		size, err := file.Read(buffer)

// 		if err != nil {
// 			break
// 		}

// 		sub_buffer := append(leftover, buffer[:size]...)

// 		for len(sub_buffer) > 0 {
// 			r, rsize := utf8.DecodeRune(sub_buffer)

// 			if r == utf8.RuneError {
// 				break
// 			}

// 			sub_buffer = sub_buffer[rsize:]

// 			if !unicode.IsSpace(r) && !is_inside_word {
// 				word_count++
// 			}

// 			is_inside_word = !unicode.IsSpace(r)
// 		}

// 		leftover = append(leftover, sub_buffer...)
// 		leftover = nil
// 	}
// 	return word_count
// }
