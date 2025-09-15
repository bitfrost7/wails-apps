package backend

import (
	"excel-tools/backend/excel"
	"fmt"
)

type WordFreqStatConfig struct {
	WfInputDir string

	IntervalNumber int
	SplitChar      string
}

// 词频统计
func (c *WordFreqStatConfig) WordFreqStat() error {
	files := excel.ReadFromDir(c.WfInputDir)
	for _, file := range files {
		for _, sheet := range file.Sheets {
			err := c.ProcessWordFreqStat(sheet.Data)
			if err != nil {
				return err
			}
			return excel.UpdateToExcel(file.FilePath+"/"+file.FileName, sheet.Data)
		}
	}
	return nil
}

func (c *WordFreqStatConfig) ProcessWordFreqStat(excelData [][]*excel.CellData) error {
	// 优先按间隔数量去统计
	if c.IntervalNumber != 0 {
		c.SplitChar = ""
	}
	for r, row := range excelData {
		var pre []string
		for i, cell := range row {
			if c.IntervalNumber != 0 {
				if i != 0 && i%c.IntervalNumber == 0 {
					stats := wordFreqStat(pre, func(s1 string, s2 string) bool {
						return s1 < s2
					})
					excelData[r][i].Value = wordFreqSum(stats)
					pre = []string{}
				} else {
					pre = append(pre, singleCh(cell.Value)...)
				}
			}
			if c.SplitChar != "" && cell.Value == c.SplitChar {
				stats := wordFreqStat(pre, func(s1 string, s2 string) bool {
					return s1 < s2
				})
				excelData[r][i].Value = wordFreqSum(stats)
				pre = []string{}
			} else {
				pre = append(pre, singleCh(cell.Value)...)
			}
		}
	}
	return nil
}

func wordFreqSum(freqs []*WordFreq) string {
	res := ""
	for _, freq := range freqs {
		res += fmt.Sprintf("%s%d", freq.word, freq.freq)
	}
	return res
}
