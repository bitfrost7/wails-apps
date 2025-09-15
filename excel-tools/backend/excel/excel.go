package excel

import (
	"archive/zip"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"unicode/utf8"
)

type XLSXFile struct {
	FilePath string
	FileName string
	Z        *zip.ReadCloser
	ColorsM  map[int]int
	Sheets   []*Sheet
}

func ReadFromDir(dir string) []*XLSXFile {
	var (
		files []*XLSXFile
		err   error
	)

	err = filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() && strings.HasSuffix(strings.ToLower(info.Name()), ".xlsx") {
			r, err := zip.OpenReader(path)
			if err != nil {
				return err
			}
			files = append(files, &XLSXFile{
				FilePath: path,
				FileName: info.Name(),
				Z:        r,
			})
		}
		return nil
	})
	if err != nil {
		panic(err)
	}
	for _, file := range files {
		file.ColorsM, err = ParseBgColor(file)
		if err != nil {
			panic(err)
			return nil
		}
		file.Sheets, err = ParseWorksheets(file)
		if err != nil {
			panic(err)
			return nil
		}
		PrintXLSXFileCN(file)
	}
	return files
}

// 颜色转中文
func colorNameCN(color int) string {
	switch color {
	case ColorRed:
		return "红"
	case ColorGreen:
		return "绿"
	case ColorBlue:
		return "蓝"
	case ColorYellow:
		return "黄"
	case ColorLightYellow:
		return "淡黄"
	default:
		return "无"
	}
}

// 打印 XLSX 文件（中文颜色 + 列宽自适应）
func PrintXLSXFileCN(xlsx *XLSXFile) {
	for _, sheet := range xlsx.Sheets {
		fmt.Printf("Sheet: %s\n", sheet.SheetName)
		printSheetCN(sheet)
		fmt.Println()
	}
}

func printSheetCN(sheet *Sheet) {
	if len(sheet.Data) == 0 {
		return
	}

	// 先计算每列最大宽度
	colWidths := make([]int, len(sheet.Data[0]))
	for _, row := range sheet.Data {
		for colIdx, cell := range row {
			content := ""
			if cell != nil {
				content = fmt.Sprintf("%s(%s)", cell.Value, colorNameCN(cell.Color))
			}
			w := utf8.RuneCountInString(content)
			if w > colWidths[colIdx] {
				colWidths[colIdx] = w
			}
		}
	}

	// 打印每行
	for rowIdx, row := range sheet.Data {
		fmt.Printf("%3d |", rowIdx+1)
		for colIdx, cell := range row {
			content := ""
			if cell != nil {
				content = fmt.Sprintf("%s(%s)", cell.Value, colorNameCN(cell.Color))
			}
			width := colWidths[colIdx]
			fmt.Printf("%-*s |", width, content)
		}
		fmt.Println()
	}
}
