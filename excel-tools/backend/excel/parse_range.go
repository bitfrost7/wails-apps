package excel

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// SqrefRange 表示Excel区域，所有索引从0开始
type SqrefRange struct {
	startRow int
	startCol int
	endRow   int
	endCol   int
}

// ParseSqref 解析Excel区域引用字符串，返回SqrefRange
// totalRows 和 totalCols 表示表格的总行数和总列数（0-based的最大索引+1）
func ParseSqref(sqref string, totalRows, totalCols int) (SqrefRange, error) {
	if totalRows <= 0 || totalCols <= 0 {
		return SqrefRange{}, fmt.Errorf("totalRows and totalCols must be positive")
	}

	// 处理空字符串或无效输入
	if strings.TrimSpace(sqref) == "" {
		return SqrefRange{}, fmt.Errorf("empty sqref string")
	}

	parts := strings.Split(sqref, ":")
	if len(parts) != 2 {
		return SqrefRange{}, fmt.Errorf("invalid sqref format: %s", sqref)
	}

	startPart := strings.TrimSpace(parts[0])
	endPart := strings.TrimSpace(parts[1])

	var startRow, startCol, endRow, endCol int
	var err error

	// 解析起始部分
	startRow, startCol, err = parseCellReference(startPart, totalRows, totalCols)
	if err != nil {
		return SqrefRange{}, err
	}

	// 解析结束部分
	endRow, endCol, err = parseCellReference(endPart, totalRows, totalCols)
	if err != nil {
		return SqrefRange{}, err
	}

	// 确保索引在有效范围内
	startRow = clamp(startRow, 0, totalRows-1)
	startCol = clamp(startCol, 0, totalCols-1)
	endRow = clamp(endRow, 0, totalRows-1)
	endCol = clamp(endCol, 0, totalCols-1)

	// 确保起始位置不大于结束位置
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}

	return SqrefRange{
		startRow: startRow,
		startCol: startCol,
		endRow:   endRow,
		endCol:   endCol,
	}, nil
}

// parseCellReference 解析单个单元格引用或行列引用（支持$符号）
func parseCellReference(ref string, totalRows, totalCols int) (int, int, error) {
	// 移除$符号
	cleanRef := strings.ReplaceAll(ref, "$", "")

	// 匹配纯列引用 (如 "A", "ET", "F", "$A", "$F")
	if isColumnOnly(cleanRef) {
		col, err := columnLetterToIndex(cleanRef)
		if err != nil {
			return 0, 0, err
		}
		// 纯列引用：从第0行到最后一行
		return 0, col, nil
	}

	// 匹配纯行引用 (如 "6", "8", "$6", "$8")
	if isRowOnly(cleanRef) {
		row, err := strconv.Atoi(cleanRef)
		if err != nil {
			return 0, 0, fmt.Errorf("invalid row reference: %s", ref)
		}
		// 纯行引用：从第0列到最后一列
		return row - 1, 0, nil
	}

	// 匹配标准单元格引用 (如 "A1", "ET68", "$A$1", "$ET$68")
	re := regexp.MustCompile(`^(\$?[A-Za-z]+\$?)(\$?\d+)$`)
	matches := re.FindStringSubmatch(ref)
	if matches != nil {
		// 提取列字母部分（移除$符号）
		colPart := strings.ReplaceAll(matches[1], "$", "")
		// 提取行数字部分（移除$符号）
		rowPart := strings.ReplaceAll(matches[2], "$", "")

		col, err := columnLetterToIndex(colPart)
		if err != nil {
			return 0, 0, err
		}

		row, err := strconv.Atoi(rowPart)
		if err != nil {
			return 0, 0, fmt.Errorf("invalid row number: %s", rowPart)
		}

		return row - 1, col, nil
	}

	return 0, 0, fmt.Errorf("invalid cell reference: %s", ref)
}

// isColumnOnly 检查是否为纯列引用
func isColumnOnly(ref string) bool {
	re := regexp.MustCompile(`^[A-Za-z]+$`)
	return re.MatchString(ref)
}

// isRowOnly 检查是否为纯行引用
func isRowOnly(ref string) bool {
	re := regexp.MustCompile(`^\d+$`)
	return re.MatchString(ref)
}

// columnLetterToIndex 将Excel列字母转换为0-based索引
func columnLetterToIndex(letters string) (int, error) {
	letters = strings.ToUpper(letters)
	result := 0

	for _, char := range letters {
		if char < 'A' || char > 'Z' {
			return 0, fmt.Errorf("invalid column letter: %s", letters)
		}
		result = result*26 + int(char-'A')
	}

	return result, nil
}

// clamp 确保值在最小值和最大值之间
func clamp(value, min, max int) int {
	if value < min {
		return min
	}
	if value > max {
		return max
	}
	return value
}
