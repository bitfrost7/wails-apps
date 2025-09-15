package backend

import (
	"errors"
	"excel-tools/backend/excel"
	"fmt"
	"math"
	"sort"
)

const (
	ColumnMode = "列统计"
	RowMode    = "行统计"

	ColorNone   = 0
	ColorRed    = 1
	ColorGreen  = 2
	ColorBlue   = 3
	ColorYellow = 4
)

var textGroup = map[int][]string{
	1: {"生", "旺", "库"},
	2: {"红", "黄", "蓝"},
	3: {"天", "日", "寅"},
	4: {"木", "上", "已"},
	5: {"地", "月", "午"},
	6: {"火", "中", "酉"},
	7: {"人", "星", "戌"},
	8: {"土", "下", "丑"},
}

var TextColorMap = map[string]int{
	"生": ColorRed, "旺": ColorGreen, "库": ColorBlue,
	"红": ColorRed, "黄": ColorGreen, "蓝": ColorBlue,
	"天": ColorRed, "日": ColorGreen, "寅": ColorBlue,
	"木": ColorRed, "上": ColorGreen, "已": ColorBlue,
	"地": ColorRed, "月": ColorGreen, "午": ColorBlue,
	"火": ColorRed, "中": ColorGreen, "酉": ColorBlue,
	"人": ColorRed, "星": ColorGreen, "戌": ColorBlue,
	"土": ColorRed, "下": ColorGreen, "丑": ColorBlue,
}

var TextWeightMap = map[string]int{
	"生": 1, "旺": 2, "库": 3,
	"红": 1, "黄": 2, "蓝": 3,
	"天": 1, "日": 2, "寅": 3,
	"木": 1, "上": 2, "已": 3,
	"地": 1, "月": 3, "午": 3,
	"火": 1, "中": 2, "酉": 3,
	"人": 1, "星": 2, "戌": 3,
	"土": 1, "下": 2, "丑": 3,
}

var TextGroupMap = map[string]int{
	"生": 1, "旺": 1, "库": 1,
	"红": 2, "黄": 2, "蓝": 2,
	"天": 3, "日": 3, "寅": 3,
	"木": 4, "上": 4, "已": 4,
	"地": 5, "月": 3, "午": 5,
	"火": 6, "中": 6, "酉": 6,
	"人": 7, "星": 7, "戌": 7,
	"土": 8, "下": 8, "丑": 8,
}

// 关键词统计

type KeyWordStatConfig struct {
	KwInputDir    string
	KwOutputDir   string
	StatMode      string
	TargetNumber  int
	ForwardNumber int
	SelectedColor string
}

type KeywordGroup struct {
	group     int
	statgroup []*KeywordFreqStat
}

type KeywordFreqStat struct {
	repeat int
	freq   []*WordFreq
}

type WordFreq struct {
	word string
	freq int
}

func sortByTextWeight(s1, s2 string) bool {
	return TextWeightMap[s1] < TextWeightMap[s2]
}

func (c *KeyWordStatConfig) KeyWordStat() error {
	files := excel.ReadFromDir(c.KwInputDir)
	for _, file := range files {
		for _, sheet := range file.Sheets {
			group, err := c.ProcessKeyWordStat(sheet.Data)
			if err != nil {
				return err
			}
			err = c.WriteKeywordGroupToXlsx(file.FileName, group)
			if err != nil {
				return err
			}
		}
	}
	return nil
}

func (c *KeyWordStatConfig) ProcessKeyWordStat(data [][]*excel.CellData) ([]*KeywordGroup, error) {
	lookBackCount := c.ForwardNumber
	targetColor := excel.ParseColor(c.SelectedColor)

	// 存储重复颜色格子数 对应的单词
	serialM := make(map[int][]string)

	// 列统计模式 从选中列开始往前遍历
	if c.StatMode == ColumnMode {
		selectedCol := c.TargetNumber
		for _, row := range data {
			if len(row) <= selectedCol {
				return nil, errors.New("选择列超过表格大小")
			}
			if selectedCol-lookBackCount <= 0 {
				return nil, errors.New("往回列超过表格大小")
			}
			repeat := 0
			for i := selectedCol - 2; i >= selectedCol-lookBackCount-1; i-- {
				if row[i].Color != targetColor {
					break
				}
				repeat++
			}
			if repeat == 0 {
				continue
			}
			serialM[repeat] = append(serialM[repeat], singleCh(row[selectedCol-1].Value)...)
		}
	}

	// 行统计模式 从选中行往上遍历
	if c.StatMode == RowMode {
		selectedRow := c.TargetNumber
		if selectedRow > len(data) {
			return nil, errors.New("选中行不能超过表格大小")
		}
		if selectedRow-lookBackCount <= 0 {
			return nil, errors.New("往回列超过表格大小")
		}
		for col := 0; col < len(data[0]); col++ {
			repeat := 0
			for i := selectedRow - 2; i >= selectedRow-lookBackCount-1; i-- {
				if data[i][col].Color != targetColor {
					break
				}
				repeat++
			}
			if repeat == 0 {
				continue
			}
			serialM[repeat] = append(serialM[repeat], singleCh(data[selectedRow-1][col].Value)...)
		}
	}

	gM := make(map[int]map[int][]*WordFreq)
	for rpt, words := range serialM {
		wordFreq := wordFreqStat(words, sortByTextWeight)
		for _, freq := range wordFreq {
			g := TextGroupMap[freq.word]
			if _, ok := gM[g]; !ok {
				gM[g] = make(map[int][]*WordFreq)
			}
			gM[g][rpt] = append(gM[g][rpt], freq)
		}
	}
	groups := make([]*KeywordGroup, 0)
	for gid, g := range gM {
		stat := make([]*KeywordFreqStat, 0)
		for rpt, freqs := range g {
			stat = append(stat, &KeywordFreqStat{
				repeat: rpt,
				freq:   freqs,
			})
		}
		sort.Slice(stat, func(i, j int) bool {
			return stat[i].repeat < stat[j].repeat
		})
		groups = append(groups, &KeywordGroup{
			group:     gid,
			statgroup: stat,
		})
	}
	return groups, nil
}

func (c *KeyWordStatConfig) WriteKeywordGroupToXlsx(fileName string, group []*KeywordGroup) error {
	maxLine := 0
	for _, g := range group {
		if maxLine < len(g.statgroup) {
			maxLine = len(g.statgroup)
		}
	}
	data := make([][]*excel.CellData, maxLine+2)
	for _, g := range group {
		// 确认表头
		header := textGroup[g.group]
		data[0] = append(data[0], &excel.CellData{
			Value: fmt.Sprintf("列-%d统计", c.TargetNumber),
		})
		for _, h := range header {
			data[0] = append(data[0], &excel.CellData{
				Value: h,
				Color: TextColorMap[h],
			})
		}
		data[0] = append(data[0], &excel.CellData{})
		for i, stat := range g.statgroup {
			data[i+1] = append(data[i+1], ProcessKeywordStatToCellData(stat)...)
		}
		// 添加计算行
		freq := make([]int, len(g.statgroup[0].freq))
		for i := len(g.statgroup) - 1; i >= 0; i-- {
			if sum := FreqSum(g.statgroup[i].freq); sum != 0 {
				for j, f := range g.statgroup[i].freq {
					freq[j] += f.freq
				}
			}
		}
		data[len(g.statgroup)+1] = append(data[len(g.statgroup)+1], &excel.CellData{
			IntValue: -1,
		})
		for _, f := range freq {
			data[len(g.statgroup)+1] = append(data[len(g.statgroup)+1], &excel.CellData{
				IntValue: f,
			})
		}
		data[len(g.statgroup)+1] = append(data[len(g.statgroup)+1], &excel.CellData{
			IntValue: -1,
		})
	}
	fileName = fmt.Sprintf("%s-%d-%d-%s.xlsx", fileName, c.TargetNumber, c.ForwardNumber, c.SelectedColor)
	return excel.SaveToExcel(c.KwOutputDir+"/"+fileName, data)
}

func ProcessKeywordStatToCellData(stat *KeywordFreqStat) []*excel.CellData {
	res := make([]*excel.CellData, 0)
	res = append(res, &excel.CellData{
		IntValue: stat.repeat,
	})
	maxInt := math.MaxInt
	for _, f := range stat.freq {
		if f.freq < maxInt {
			maxInt = f.freq
		}
	}
	for _, f := range stat.freq {
		cell := &excel.CellData{
			IntValue: f.freq,
		}
		if f.freq > maxInt {
			cell.Color = ColorYellow
		}
		res = append(res, cell)
	}
	res = append(res, &excel.CellData{})
	return res
}
