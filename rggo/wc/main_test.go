package main

import (
	"bytes"
	"testing"
)

func TestCountWords(t *testing.T) {
	var b = bytes.NewBufferString("computer language science keyboard")
	var exp = 4
	res, err := count(b, false)
	if err != nil {
		t.Fatalf("Unexpected error: %v\n", err)
	}

	if res != exp {
		t.Errorf("Expected %d, got %d instead.\n", exp, res)
	}
}

func TestCountLines(t *testing.T) {
	var b = bytes.NewBufferString("Are you OK?\n I'm fine\nFuck you")
	var exp = 3

	res, err := count(b, true)
	if err != nil {
		t.Fatalf("Unexpected error: %v\n", err)
	}
	if res != exp {
		t.Errorf("Expected %d, got %d instead.\n", exp, res)
	}
}
