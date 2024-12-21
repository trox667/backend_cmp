package service

import (
	"de.trox667/backend/model"
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
)

func toInt64(value string, defaultValue int64) int64 {
	val, err :=strconv.ParseInt(value, 10, 64)
	if err != nil {
		return defaultValue
	}
	return val
}

func toFloat64(value string, defaultValue float64) float64 {
	val, err := strconv.ParseFloat(value, 64)
	if err != nil {
		return defaultValue
	}
	return val
}

func ReadExcel() []model.Entry {
	reader, err := excelize.OpenFile("sample.xlsx")
	if err != nil {
		fmt.Println(err)
		return nil
	}

	defer func() {
		if err := reader.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	activeSheet := reader.GetActiveSheetIndex()
	sheetName := reader.GetSheetName(activeSheet)
	rows, err := reader.GetRows(sheetName)
	entries := make([]model.Entry, len(rows))
	for i, row := range rows[3:] {
		entry := model.Entry{}
		entry.Id = toInt64(strings.TrimSpace(row[0]), 0)
		entry.Title = strings.TrimSpace(row[1])
		entry.Category = strings.TrimSpace(row[2])
		entry.Income = toFloat64(strings.TrimSpace(row[3]), 0.0)
		entry.Stock = toInt64(strings.TrimSpace(row[4]), 0)
		entries[i] = entry
	}
	return entries
}
