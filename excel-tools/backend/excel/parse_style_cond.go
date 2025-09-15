package excel

import (
	"errors"
	"sort"
	"strconv"
	"strings"
)

// 条件格式相关的结构体
type ConditionalFormatting struct {
	Sqref  string   `xml:"sqref,attr"`
	CfRule []CfRule `xml:"cfRule"`
}

type CfRule struct {
	Type     string   `xml:"type,attr"`
	DxfId    int      `xml:"dxfId,attr"`
	Priority int      `xml:"priority,attr"`
	Operator string   `xml:"operator,attr,omitempty"`
	Text     string   `xml:"text,attr,omitempty"`
	Formula  []string `xml:"formula"`
}

// parseCellStyle 解析条件格式并应用到单元格数据
func parseCellStyle(data [][]*CellData, cond []ConditionalFormatting, colorsM map[int]int) {
	if cond == nil {
		return
	}
	var row, col int
	row = len(data)
	if row == 0 {
		col = 0
	} else {
		col = len(data[0])
	}

	// 先排序
	for _, format := range cond {
		if len(format.CfRule) == 0 {
			panic("invalid format")
		}
		sort.Slice(format.CfRule, func(i, j int) bool {
			return format.CfRule[i].Priority > format.CfRule[j].Priority
		})
	}
	sort.Slice(cond, func(i, j int) bool {
		return cond[i].CfRule[0].Priority > cond[j].CfRule[0].Priority
	})

	for _, format := range cond {
		// 解析范围引用（如"A1:U20"）
		ranges, err := parseRangeRef(format.Sqref, row, col)
		if err != nil {
			continue
		}

		// 应用每个条件格式规则
		for _, rule := range format.CfRule {
			applyCfRuleToRange(data, rule, colorsM, ranges)
		}
	}
}

// 解析多段范围引用，如 "$A$1:$ET$2,$14:$14,$19:$20"
func parseRangeRef(rangeRef string, row, col int) ([]*SqrefRange, error) {
	var ranges []*SqrefRange
	parts := strings.Split(rangeRef, " ")
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		if rg, err := ParseSqref(part, row, col); err == nil {
			ranges = append(ranges, &rg)
		}
	}
	return ranges, nil
}

// 解析单个单元格引用，如 $A$1、B2
func parseSingleCellRef(cellRef string) (row, col int, err error) {
	cellRef = strings.ReplaceAll(cellRef, "$", "")
	var colStr, rowStr string
	for i, r := range cellRef {
		if r >= '0' && r <= '9' {
			colStr = cellRef[:i]
			rowStr = cellRef[i:]
			break
		}
	}
	if colStr == "" || rowStr == "" {
		return 0, 0, errors.New("invalid cell ref: " + cellRef)
	}

	// 列字母转数字，索引从 0 开始
	col = 0
	for _, c := range colStr {
		col = col*26 + int(c-'A')
	}

	// 行号转索引，从 0 开始
	r, err := strconv.Atoi(rowStr)
	if err != nil {
		return 0, 0, err
	}
	row = r - 1

	return row, col, nil
}

func applyCfRuleToRange(cellData [][]*CellData, rule CfRule, colorsM map[int]int, ranges []*SqrefRange) {
	for _, r := range ranges {
		// 边界裁剪，防止越界
		startRow := r.startRow
		if startRow < 0 {
			startRow = 0
		}
		endRow := r.endRow
		if endRow >= len(cellData) {
			endRow = len(cellData) - 1
		}

		for row := startRow; row <= endRow; row++ {
			if cellData[row] == nil {
				continue
			}
			startCol := r.startCol
			if startCol < 0 {
				startCol = 0
			}
			endCol := r.endCol
			if endCol >= len(cellData[row]) {
				endCol = len(cellData[row]) - 1
			}

			for col := startCol; col <= endCol; col++ {
				cell := cellData[row][col]
				if cell != nil && checkCondition(cell, rule) {
					if color, ok := colorsM[rule.DxfId]; ok {
						cell.Color = color
					}
				}
			}
		}
	}
}

// 检查条件是否满足
func checkCondition(cell *CellData, rule CfRule) bool {
	switch rule.Type {
	case "notContainsBlanks":
		// 检查不为空
		return strings.TrimSpace(cell.Value) != ""

	case "containsText":
		// 检查包含特定文本
		if rule.Text != "" {
			return strings.Contains(cell.Value, rule.Text)
		}
		return false

	default:
		return false
	}
}
