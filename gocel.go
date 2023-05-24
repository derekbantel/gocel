package gocel

import "github.com/derekbantel/gocel/internal"

func ReadExcel(workbookLocation string, sheetName string) [][]string {
	internal.ExtractXMLFromExcel(workbookLocation)
	stringMap := internal.ParseSharedStrings()
	sheetData := internal.ParseSheetData(sheetName)
	newData := internal.CombineStrWData(stringMap, sheetData)
	newArray := internal.CreateTwoDArray(newData)

	return newArray
}
