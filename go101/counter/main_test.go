package main_test

import (
	counter "counter"
	"strings"
	"testing"
)

func TestCountWords(t *testing.T) {
	// type TestCase struct {
	// 	name     string
	// 	input    string
	// 	expected int
	// }

	test_cases := []struct {
		name     string
		input    string
		expected int
	}{
		{
			name:     "empty",
			input:    "",
			expected: 0,
		},
		{
			name:     "multispace",
			input:    "          ",
			expected: 0,
		},
		{
			name:     "single word",
			input:    "hello",
			expected: 1,
		},
		{
			name:     "multiple words",
			input:    "hello world",
			expected: 2,
		},
		{
			name:     "multiple words with tabs",
			input:    "hello\tworld",
			expected: 2,
		},
		{
			name:     "multiple words with newlines",
			input:    "hello\nworld",
			expected: 2,
		},
		{
			name:     "multiple words with multiple spaces",
			input:    "hello  world",
			expected: 2,
		},
		{
			name:     "multiple words with leading spaces",
			input:    " hello world",
			expected: 2,
		},
		{
			name:     "multiple words with trailing spaces",
			input:    "hello world  ",
			expected: 2,
		},
		{
			name:     "multiple words with leading and trailing spaces",
			input:    " hello hello world  ",
			expected: 3,
		},
	}
	for _, test_case := range test_cases {
		t.Run(test_case.name, func(t *testing.T) {
			reader := strings.NewReader(test_case.input)
			result := counter.CountWords(reader)
			if result != test_case.expected {
				t.Logf("%s: %d != %d", test_case.name, result, test_case.expected)
				t.Fail()
			}
		})
	}
}
