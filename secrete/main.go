package main

import (
	"fmt"
	"strings"
	"flag"
	"github.com/alexellis/hmac"
)

func main() {
	var input string
	var secret string

	flag.StringVar(&input, "message","", "message to create a digest from")
	flag.StringVar(&secret, "secret", "", "secret for the digest")
	flag.Parse() 

	if len(strings.TrimSpace(secret)) == 0 {
		panic("--secret is required")
	}

	var digest = hmac.Sign([]byte(input), []byte(secret))
	fmt.Printf("%x", digest)
}