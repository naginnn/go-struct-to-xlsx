package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
	"strconv"
	"unicode/utf8"
)

func CombineMap(first, second map[string][]any) map[string][]any {
	combined := map[string][]any{}
	for k, v := range first {
		for _, val := range v {
			combined[k] = append(combined[k], val)
		}

	}
	for k, v := range second {
		for _, val := range v {
			combined[k] = append(combined[k], val)
		}
	}
	return combined
}

func getWord(i int) string {
	if i > 26 || i < 0 {
		i = 1
	}
	return string(rune('A' - 1 + i))
}

func GetFieldAxis(t any) (int, int) {
	var x, y int
	arr := reflect.ValueOf(t)
	y = arr.Len()
	s := reflect.ValueOf(t)
	for i := 0; i < arr.Len(); i++ {
		z := s.Index(i).Interface()
		r := reflect.ValueOf(z)
		if x == 0 {
			x = r.NumField()
		}
		for j := 0; j < r.NumField(); j++ {
			if reflect.TypeOf(r.Field(j).Interface()).Kind() == reflect.Slice {
				xx, yy := GetFieldAxis(r.Field(j).Interface())
				y = y + yy
				x = xx + x - 1
			}
		}
		break
	}
	return x, y
}

func AutoFitColumns(f *excelize.File, sheetName string) error {
	cols, err := f.GetCols(sheetName)
	if err != nil {
		return err
	}
	for idx, col := range cols {
		largestWidth := 0
		for _, rowCell := range col {
			cellWidth := utf8.RuneCountInString(rowCell) + 2 // + 2 for margin
			if cellWidth > largestWidth {
				largestWidth = cellWidth
			}
		}
		name, err := excelize.ColumnNumberToName(idx + 1)
		if err != nil {
			return err
		}
		err = f.SetColWidth(sheetName, name, name, float64(largestWidth))
		if err != nil {
			return err
		}
	}
	return nil
}

func GetMap(t any, fields map[string][]any) (f map[string][]any) {
	s := reflect.ValueOf(t)
	for i := 0; i < s.Len(); i++ {
		z := s.Index(i).Interface()
		r := reflect.ValueOf(z)
		iType := r.Type()
		for j := 0; j < r.NumField(); j++ {
			if reflect.TypeOf(r.Field(j).Interface()).Kind() == reflect.Slice {
				//repeat = r.Field(j).Len()
				l := r.Field(j).Len() - 1
				for rp := 0; rp < l; rp++ {
					for d := 0; d < r.NumField(); d++ {
						if reflect.TypeOf(r.Field(d).Interface()).Kind() != reflect.Slice {
							fields[iType.Field(d).Name] = append(fields[iType.Field(d).Name], r.Field(d).Interface())
						}

					}
				}
				newFields := GetMap(r.Field(j).Interface(), make(map[string][]any))
				fields = CombineMap(fields, newFields)
			} else {
				fields[iType.Field(j).Name] = append(fields[iType.Field(j).Name], r.Field(j).Interface())
			}

		}
	}
	return fields
}

func MakeSheetFMap(xlsx *excelize.File, fields map[string][]any, headers map[string]string, sheetName string,
	autoFit, autoFilter bool) error {
	index, _ := xlsx.NewSheet(sheetName)
	xlsx.SetActiveSheet(index)
	i := 1
	for header, cols := range fields {
		if headers != nil {
			if headers[header] != "" {
				header = headers[header]
			}
		}
		err := xlsx.SetCellValue(sheetName, getWord(i)+"1", header)
		if err != nil {
			return err
		}
		for j, col := range cols {
			err := xlsx.SetCellValue(sheetName, getWord(i)+strconv.Itoa(j+2), col)
			if err != nil {
				return err
			}
		}
		i++
	}
	if autoFit {
		err := AutoFitColumns(xlsx, sheetName)
		if err != nil {
			return err
		}
	}
	if autoFilter {
		err := xlsx.AutoFilter(sheetName, fmt.Sprintf("A1:%s1", getWord(len(fields))), []excelize.AutoFilterOptions{})
		if err != nil {
			return err
		}
	}
	return nil
}
