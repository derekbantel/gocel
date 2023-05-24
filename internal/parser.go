package internal

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

type Sst struct {
	XMLName xml.Name `xml:"sst"`
	Si      []Si     `xml:"si"`
}

type Si struct {
	T string `xml:"t"`
}

type Worksheet struct {
	XMLName   xml.Name  `xml:"worksheet"`
	SheetData SheetData `xml:"sheetData"`
}

type SheetData struct {
	Rows []Row `xml:"row"`
}

type Row struct {
	R string `xml:"r,attr"`
	C []C    `xml:"c"`
}

type C struct {
	R string `xml:"r,attr"`
	T string `xml:"t,attr"`
	V string `xml:"v"`
}

func CreateTwoDArray(data Worksheet) [][]string {
	newArray := [][]string{}
	for _, row := range data.SheetData.Rows {
		tempArray := []string{}
		for i := range row.C {
			tempArray = append(tempArray, row.C[i].V)
		}

		newArray = append(newArray, tempArray)
	}

	return newArray
}

func CombineStrWData(stringMap map[int]string, data Worksheet) Worksheet {
	for _, row := range data.SheetData.Rows {
		for i := range row.C {
			if row.C[i].T == "s" {
				tempCellValue, err := strconv.Atoi(strings.TrimSpace(row.C[i].V))
				if err != nil {
					fmt.Println("could not convert string to int")
				}
				row.C[i].V = stringMap[tempCellValue]
			}
		}
	}

	return data
}

func ParseSheetData(sheetName string) Worksheet {
	worksheet := Worksheet{}

	xmlFile, err := os.Open(fmt.Sprintf("./datafiles/xl/worksheets/%v.xml", sheetName))
	if err != nil {
		fmt.Println("Error opening file:", err)
		return Worksheet{}
	}
	defer xmlFile.Close()

	byteValue, _ := io.ReadAll(xmlFile)

	xml.Unmarshal(byteValue, &worksheet)

	return worksheet
}

func ParseSharedStrings() map[int]string {
	stringMap := make(map[int]string)
	decodedData := Sst{}

	xmlFile, err := os.Open("./datafiles/xl/sharedStrings.xml")
	if err != nil {
		fmt.Println("could not open sharedstring.xml")
	}
	defer xmlFile.Close()

	byteValue, _ := io.ReadAll(xmlFile)

	xml.Unmarshal(byteValue, &decodedData)

	for i := 0; i < len(decodedData.Si); i++ {
		stringMap[i] = decodedData.Si[i].T
	}

	return stringMap
}

func ExtractXMLFromExcel(worksheetLocation string) {
	r, err := zip.OpenReader(worksheetLocation)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer r.Close()

	for _, f := range r.File {
		rc, err := f.Open()
		if err != nil {
			fmt.Println(err)
			return
		}

		path := filepath.Join("./datafiles/", f.Name)

		if f.FileInfo().IsDir() {
			os.MkdirAll(path, os.ModePerm)
		} else {
			var dir string
			if lastIndex := strings.LastIndex(path, string(os.PathSeparator)); lastIndex > -1 {
				dir = path[:lastIndex]
				os.MkdirAll(dir, os.ModePerm)
			}
		}

		outFile, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, f.Mode())
		if err != nil {
			fmt.Println(err)
			return
		}

		_, err = io.Copy(outFile, rc)
		if err != nil {
			fmt.Println(err)
		}

		outFile.Close()
		rc.Close()
	}
}
