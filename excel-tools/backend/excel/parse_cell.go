package excel

import (
	"encoding/xml"
	"io"
	"strconv"
	"strings"
)

type Sheet struct {
	SheetName string
	Data      [][]*CellData
}

type CellData struct {
	Value    string
	IntValue int
	Color    int
}

type SharedStrings struct {
	SI []struct {
		T string `xml:"t"`
	} `xml:"si"`
}

type SheetXML struct {
	XMLName   xml.Name `xml:"worksheet"`
	SheetData struct {
		Rows []struct {
			R  int `xml:"r,attr"`
			Cs []struct {
				R string `xml:"r,attr"` // A1,B2...
				T string `xml:"t,attr"` // s=sharedString
				V string `xml:"v"`      // 索引或值
				S int    `xml:"s,attr"` // 样式索引
			} `xml:"c"`
		} `xml:"row"`
	} `xml:"sheetData"`
	ConditionalFormattings []ConditionalFormatting `xml:"conditionalFormatting"`
}

type Workbook struct {
	Sheets struct {
		Sheet []struct {
			Name    string `xml:"name,attr"`
			SheetID string `xml:"sheetId,attr"`
			ID      string `xml:"id,attr"` // relationship id
		} `xml:"sheet"`
	} `xml:"sheets"`
}

func ParseWorksheets(xlsx *XLSXFile) ([]*Sheet, error) {
	// 读取 sharedStrings.xml
	var sharedStrings []string
	for _, f := range xlsx.Z.File {
		if strings.HasSuffix(f.Name, "sharedStrings.xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			var sst SharedStrings
			if err := xml.Unmarshal(data, &sst); err != nil {
				return nil, err
			}
			for _, si := range sst.SI {
				sharedStrings = append(sharedStrings, si.T)
			}
			rc.Close()
			break
		}
	}

	var workbook Workbook
	var sheetFiles = make(map[string]string) // r:id -> sheet*.xml

	for _, f := range xlsx.Z.File {
		if strings.HasSuffix(f.Name, "workbook.xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			if err := xml.Unmarshal(data, &workbook); err != nil {
				return nil, err
			}
			rc.Close()
		}
		if strings.HasSuffix(f.Name, "workbook.xml.rels") {
			rc, _ := f.Open()
			type Relationships struct {
				Relationship []struct {
					Id     string `xml:"Id,attr"`
					Target string `xml:"Target,attr"`
				} `xml:"Relationship"`
			}
			var rels Relationships
			data, _ := io.ReadAll(rc)
			if err := xml.Unmarshal(data, &rels); err != nil {
				return nil, err
			}
			for _, r := range rels.Relationship {
				sheetFiles[r.Id] = r.Target
			}
			rc.Close()
		}
	}

	// 遍历每个 sheet
	var sheets []*Sheet
	for _, s := range workbook.Sheets.Sheet {
		target := sheetFiles[s.ID]
		var sheetXML SheetXML
		for _, f := range xlsx.Z.File {
			if strings.HasSuffix(f.Name, target) {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				if err := xml.Unmarshal(data, &sheetXML); err != nil {
					return nil, err
				}
				rc.Close()
				break
			}
		}

		sheet := &Sheet{
			SheetName: s.Name,
		}
		for _, row := range sheetXML.SheetData.Rows {
			rowData := make([]*CellData, 0)
			for _, c := range row.Cs {
				val := c.V
				if c.T == "s" {
					idx, _ := strconv.Atoi(c.V)
					if idx >= 0 && idx < len(sharedStrings) {
						val = sharedStrings[idx]
					}
				}
				rowData = append(rowData, &CellData{
					Value: val,
				})
			}
			sheet.Data = append(sheet.Data, rowData)
		}
		// 解析条件格式
		parseCellStyle(sheet.Data, sheetXML.ConditionalFormattings, xlsx.ColorsM)

		sheets = append(sheets, sheet)
	}
	return sheets, nil
}
