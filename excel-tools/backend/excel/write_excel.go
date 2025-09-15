package excel

import (
	"github.com/xuri/excelize/v2"
)

// SaveToExcel 创建新的 Excel 文件并保存数据
func SaveToExcel(filePath string, data [][]*CellData) error {
	f := excelize.NewFile()
	sheet := "Sheet1"

	if err := writeDataToSheet(f, sheet, data); err != nil {
		return err
	}

	return f.SaveAs(filePath)
}

// UpdateToExcel 更新现有的 Excel 文件
func UpdateToExcel(filePath string, data [][]*CellData) error {
	// 打开现有的 Excel 文件
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		// 如果文件不存在，创建新文件
		f = excelize.NewFile()
	}

	sheet := "Sheet1"

	// 确保工作表存在
	index, _ := f.GetSheetIndex(sheet)
	if index == -1 {
		f.NewSheet(sheet)
	}

	if err := writeDataToSheet(f, sheet, data); err != nil {
		return err
	}

	return f.Save()
}

// writeDataToSheet 通用的数据写入逻辑
func writeDataToSheet(f *excelize.File, sheet string, data [][]*CellData) error {
	for r, row := range data {
		for c, cell := range row {
			if cell == nil {
				continue
			}

			rowIndex := r + 1
			colIndex := c + 1

			axis, err := excelize.CoordinatesToCellName(colIndex, rowIndex)
			if err != nil {
				return err
			}

			// 写值：优先字符串，其次整数
			if cell.Value != "" {
				f.SetCellValue(sheet, axis, cell.Value)
			} else if cell.IntValue != -1 {
				f.SetCellInt(sheet, axis, int64(cell.IntValue))
			}

			// 应用样式
			if err := applyCellStyle(f, sheet, axis, cell); err != nil {
				return err
			}
		}
	}
	return nil
}

// applyCellStyle 应用单元格样式
func applyCellStyle(f *excelize.File, sheet, axis string, cell *CellData) error {
	if cell.Color == 0 {
		return nil
	}

	var fillColor string
	switch cell.Color {
	case ColorRed:
		fillColor = "#FF0000"
	case ColorGreen:
		fillColor = "#00FF00"
	case ColorBlue:
		fillColor = "#0000FF"
	case ColorYellow:
		fillColor = "#FFFF00"
	case ColorLightYellow:
		fillColor = "#FFFACD"
	default:
		return nil
	}

	styleID, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{fillColor},
			Pattern: 1,
		},
	})
	if err != nil {
		return err
	}

	return f.SetCellStyle(sheet, axis, axis, styleID)
}
