package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

type OutputMap struct {
	Results []Document `json:"results"`
	Error   string     `json:"error"`
}

type Document struct {
	Name         string   `json:"name"`
	Version      string   `json:"version"`
	Date         string   `json:"date"`
	Description  string   `json:"description"`
	Owner        string   `json:"owner"`
	OwnerFIO     string   `json:"ownerFIO"`
	Comments     string   `json:"comments"`
	Comments2    string   `json:"comments2"`
	RelevanceRus []string `json:"relevanceRus"`
	RelevanceEng []string `json:"relevanceEng"`
}

type SearchMap struct {
	Date  string `json:"date"`
	Owner string `json:"owner"`
	Words string `json:"words"`
}

func includes(slice []string, str string) bool {
	for _, elem := range slice {
		if elem == str {
			return true
		}
	}
	return false
}

func prepareRelevanceRus(str string, sep string, links []string) []string {
	result := make([]string, 0)
	var i int = 0
	for _, elem := range strings.Split(str, sep) {
		if elem != "" {
			if len(links) != 0 && len(links) > i {
				result = append(result, elem+"==="+links[i])
				i++
			} else {
				result = append(result, elem)
			}
		}
	}
	return result
}

func prepareRelevanceEng(str string, sep string, links []string) []string {
	result := make([]string, 0)
	var i int = 0
	for _, elem := range strings.Split(str, sep) {
		if elem != "" {
			if len(links) != 0 && len(links)-1-i >= 0 {
				result = append(result, elem+"==="+links[len(links)-1-i])
				i++
			} else {
				result = append(result, elem)
			}
		}
	}
	return result
}

func searchDocuments(documents [][]string, searcher SearchMap) [][]string {
	var searchResult [][]string
	words := strings.Split(strings.TrimSpace(strings.ToLower(searcher.Words)), " ")

	for _, row := range documents {
		var checkData, checkOwner, checkWords = false, false, true
		if searcher.Owner == "" ||
			strings.Contains(strings.ToLower(row[10]), strings.ToLower(searcher.Owner)) {
			checkOwner = true
		}
		if searcher.Date == "" || row[3] == searcher.Date {
			checkData = true
		}

		for _, word := range words {
			if strings.Contains(strings.ToLower(row[1] + row[12]), word) {
				checkWords = true
				continue
			}
			checkWords = false
			break
		}

		if checkWords && checkOwner && checkData {
			searchResult = append(searchResult, [][]string{row}...)
		}
	}
	return searchResult
}

func findLinks(row []string) []string {
	var links []string
	for _, cell := range row {
		if strings.Contains(cell, "https://") {
			links = append(links, cell)
		}
	}
	return links
}

func normalizeDocSearch(searchEntries [][]string) OutputMap {
	normalizedDocument := make([]Document, 0)
	for _, row := range searchEntries {
		links := findLinks(row[16:])
		relRus := prepareRelevanceRus(row[14], " | ", links)
		relEng := prepareRelevanceEng(row[15], " | ", links)
		document := Document{
			Name:         row[1],
			Version:      row[2],
			Date:         row[3],
			Description:  row[8],
			Owner:        row[9],
			OwnerFIO:     row[10],
			Comments:     row[12],
			Comments2:    row[13],
			RelevanceRus: relRus,
			RelevanceEng: relEng,
		}
		normalizedDocument = append(normalizedDocument, document)
	}
	return OutputMap{
		Results: normalizedDocument,
		Error:   "",
	}
}

func printError(err error) {
	output := OutputMap{
		Results: make([]Document, 0),
		Error:   err.Error(),
	}
	outputBytes, err := json.Marshal(output)
	fmt.Fprintf(os.Stdout, "%s\n", string(outputBytes))
}

func main() {
	var documents [][]string
	var fileName string
	var searchString string

	flag.StringVar(&fileName, "file", "", "xls file name to read from")
	flag.StringVar(&searchString, "search", "", "json string to search in excel file")

	flag.Parse()

	// Unmarshal search string parameter to struct
	var searcher SearchMap
	err := json.Unmarshal([]byte(searchString), &searcher)
	if err != nil {
		printError(err)
		return
	}

	f, err := excelize.OpenFile(fileName)
	if err != nil {
		printError(err)
		return
	}

	// Get all rows from excel list
	rows, err := f.GetRows("ВСЕ ПРОЕКТЫ")
	if err != nil {
		printError(err)
		return
	}

	// iterate through all rows and collect which includes needed string
	for _, row := range rows {
		var col []string
		for _, colCell := range row {
			col = append(col, colCell)
		}
		if includes(col, "Опубликован") {
			documents = append(documents, [][]string{col}...)
		}
	}

	searchResult := normalizeDocSearch(searchDocuments(documents, searcher))
	res, err := json.Marshal(searchResult)
	fmt.Fprintf(os.Stdout, "%s\n", string(res))
}
