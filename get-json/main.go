package main

import (
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
)

type Person struct {
	Name  string `json:"name"`
	Craft string `json:"craft"`
}

type Astro struct {
	Number int      `json:"number"`
	People []Person `json:"people"`
}

func main() {
	apiURL := "http://api.open-notify.org/astros.json"
	people, err := getAstros(apiURL)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Printf("%d people found in space.\n", people.Number)
	for _, p := range people.People {
		fmt.Printf("Let's wave to: %s\n", p.Name)
	}
}

func getAstros(apiURL string) (Astro, error) {
	p := Astro{}
	req, err := http.NewRequest(http.MethodGet, apiURL, nil)
	if err != nil {
		return p, err
	}

	req.Header.Set("User-Agent", "spacecount-tutorial")
	res, err := http.DefaultClient.Do(req)
	if err != nil {
		return p, err
	}

	if res.Body != nil {
		defer res.Body.Close()
	}

	body, err := io.ReadAll(res.Body)
	if err != nil {
		return p, err
	}

	err = json.Unmarshal(body, &p)
	if err != nil {
		return p, err
	}
	return p, nil
}

// {
//   "people": [
//     {
//       "craft": "ISS",
//       "name": "Oleg Kononenko"
//     },
//     {
//       "craft": "ISS",
//       "name": "Nikolai Chub"
//     },
//     {
//       "craft": "ISS",
//       "name": "Tracy Caldwell Dyson"
//     },
//     {
//       "craft": "ISS",
//       "name": "Matthew Dominick"
//     },
//     {
//       "craft": "ISS",
//       "name": "Michael Barratt"
//     },
//     {
//       "craft": "ISS",
//       "name": "Jeanette Epps"
//     },
//     {
//       "craft": "ISS",
//       "name": "Alexander Grebenkin"
//     },
//     {
//       "craft": "ISS",
//       "name": "Butch Wilmore"
//     },
//     {
//       "craft": "ISS",
//       "name": "Sunita Williams"
//     },
//     {
//       "craft": "Tiangong",
//       "name": "Li Guangsu"
//     },
//     {
//       "craft": "Tiangong",
//       "name": "Li Cong"
//     },
//     {
//       "craft": "Tiangong",
//       "name": "Ye Guangfu"
//     }
//   ],
//   "number": 12,
//   "message": "success"
// }
