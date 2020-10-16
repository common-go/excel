package excel

import (
	"context"
	"errors"
	"fmt"
	. "github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
	"time"
)

func ExportMaps(ctx context.Context, headers []string, arr []map[interface{}]interface{}, format string) {
	f := NewFile()
	for i, header := range headers {
		headerCellName, _ := CoordinatesToCellName(i+1, 1)
		f.SetCellValue("Sheet1", headerCellName, header)
		for index, rowMap := range arr {
			cellName, _ := CoordinatesToCellName(i+1, index+2)
			if v, ok := rowMap[header].(time.Time); ok {
				time, err := timeToExcelTime(v.UTC())
				if err != nil {
					fmt.Println(err)
				}
				f.SetCellValue("Sheet1", cellName, time)
				style, _ := f.NewStyle(&Style{NumFmt: 22})
				f.SetCellStyle("Sheet1", cellName, cellName, style)
			} else {
				f.SetCellValue("Sheet1", cellName, rowMap[header])
			}
		}
	}
	if err := f.SaveAs(fmt.Sprintf("export_%v.%v ", strconv.Itoa(int(time.Now().Unix())), format)); err != nil {
		println(err.Error())
	}
}

const (
	dayNanoseconds = 24 * time.Hour
	maxDuration    = 290 * 364 * dayNanoseconds
)

var (
	excelMinTime1900      = time.Date(1899, time.December, 31, 0, 0, 0, 0, time.UTC)
	excelBuggyPeriodStart = time.Date(1900, time.March, 1, 0, 0, 0, 0, time.UTC).Add(-time.Nanosecond)
)

func timeToExcelTime(t time.Time) (float64, error) {
	if t.Location() != time.UTC {
		return 0.0, errors.New("only UTC time expected")
	}
	if t.Before(excelMinTime1900) {
		return 0.0, nil
	}
	tt := t
	diff := t.Sub(excelMinTime1900)
	result := float64(0)
	for diff >= maxDuration {
		result += float64(maxDuration / dayNanoseconds)
		tt = tt.Add(-maxDuration)
		diff = tt.Sub(excelMinTime1900)
	}
	rem := diff % dayNanoseconds
	result += float64(diff-rem)/float64(dayNanoseconds) + float64(rem)/float64(dayNanoseconds)
	if t.After(excelBuggyPeriodStart) {
		result += 1.0
	}
	return result, nil
}
