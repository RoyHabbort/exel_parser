package main

import (
	"encoding/xml"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io/ioutil"
	"os"
)


type Records struct {
	XMLName xml.Name `xml:"records"`
	Version int      `xml:"version,attr"`
	Record   []Record `xml:"record"`
}

type Record struct {
	Id 			string `xml:"id"`
	FirstName   string `xml:"FirstName"`
	SecondName  string `xml:"SecondName"`
	Gender 		string `xml:"Gender"`
	Country 	string `xml:"Country"`
	Age    		string `xml:"Age"`
	Date    	string `xml:"Date"`
}

func main() {

	var nameExcel string
	var nameXml string

	fmt.Print("Введите имя Excel файла: ")
	fmt.Fscan(os.Stdin, &nameExcel)

	fmt.Print("Введите имя Xml файла: ")
	fmt.Fscan(os.Stdin, &nameXml)

	fullFileNameForExcel := "./excel/" + nameExcel + ".xlsx"
	fullFileNameForXml := "./xml/" + nameXml + ".xml"

	xlsx, err := excelize.OpenFile(fullFileNameForExcel)
	if err != nil {
		fmt.Println(err)
		return
	}

	Records := Records{
		Record: []Record{},
	}

	// Get all the rows in the Sheet1.
	rows, err := xlsx.GetRows("Sheet1")
	for keyRow, row := range rows {
		if keyRow == 0 {
			continue
		}
		record := Record{}

		for keyCell, colCell := range row {
			if keyCell == 1 {
				record.FirstName = colCell
			}
			if keyCell == 2 {
				record.SecondName = colCell
			}
			if keyCell == 3 {
				record.Gender = colCell
			}
			if keyCell == 4 {
				record.Country = colCell
			}
			if keyCell == 5 {
				record.Age = colCell
			}
			if keyCell == 6 {
				record.Date = colCell
			}
			if keyCell == 7 {
				record.Id = colCell
			}
		}

		Records.Record = append(Records.Record, record)
	}

	resultXml, _ := xml.MarshalIndent(Records, "", " ")
	fmt.Print(string(resultXml))
	_ = ioutil.WriteFile(fullFileNameForXml, resultXml, 0644)
}
