package main

import (
	"de.trox667/backend/service"
	"encoding/json"
	"fmt"
	"net/http"
	"time"
)

func readExcel(w http.ResponseWriter, request *http.Request) {
	now := time.Now()
	entries := service.ReadExcel()
	then := time.Now()
	fmt.Printf("Excel parsing took %dms\n", then.Sub(now).Milliseconds())
	encoder := json.NewEncoder(w)
	err := encoder.Encode(&entries)
	if err != nil {
		http.Error(w, http.StatusText(http.StatusInternalServerError), http.StatusInternalServerError)
		return
	}
}

func main() {
	fmt.Println("Hello World")

	http.HandleFunc("/", readExcel)
	fmt.Println("Starting server on 8080")
	err := http.ListenAndServe("localhost:8080", nil)
	if err != nil {
		return
	}
}
