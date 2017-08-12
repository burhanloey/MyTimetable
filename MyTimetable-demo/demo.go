package demo

import (
	"io/ioutil"
	"net/http"
)

var index []byte

func init() {
	content, err := ioutil.ReadFile("index.html")
	if err != nil {
		panic(err)
	}
	index = content

	http.HandleFunc("/", root)
}

func root(w http.ResponseWriter, r *http.Request) {
	w.Write(index)
}
