package excel

import (
	"encoding/xml"
	"io"
	"math"
	"strconv"
)

// 代表 <dxfs> 下的 <dxf>
type Dxf struct {
	Font   *Font   `xml:"font"`
	Fill   *Fill   `xml:"fill"`
	Border *Border `xml:"border"`
}

type Font struct {
	Bold *struct {
		Val string `xml:"val,attr"`
	} `xml:"b"`
	Color *Color `xml:"color"`
}

type Color struct {
	Rgb   string `xml:"rgb,attr"`
	Theme string `xml:"theme,attr"`
	Tint  string `xml:"tint,attr"`
}

type Fill struct {
	PatternFill *PatternFill `xml:"patternFill"`
}

type PatternFill struct {
	PatternType string `xml:"patternType,attr"`
	FgColor     *Color `xml:"fgColor"`
	BgColor     *Color `xml:"bgColor"`
}

type Border struct {
	Top    *BorderStyle `xml:"top"`
	Bottom *BorderStyle `xml:"bottom"`
	Left   *BorderStyle `xml:"left"`
	Right  *BorderStyle `xml:"right"`
}

type BorderStyle struct {
	Style string `xml:"style,attr"`
	Color *Color `xml:"color"`
}

type Styles struct {
	XMLName xml.Name `xml:"styleSheet"`
	Dxfs    struct {
		Count int   `xml:"count,attr"`
		Dxfs  []Dxf `xml:"dxf"`
	} `xml:"dxfs"`
}

type RGB struct {
	R, G, B uint8
}

var (
	ColorNone        = 0
	ColorRed         = 1
	ColorGreen       = 2
	ColorBlue        = 3
	ColorYellow      = 4
	ColorLightYellow = 5

	Red    = RGB{255, 0, 0}
	Yellow = RGB{255, 255, 0}
	Blue   = RGB{0, 0, 255}
	Green  = RGB{0, 255, 0}
)

var colors = []RGB{
	Red, Yellow, Blue, Green,
}

func ColorName(color int) string {
	switch color {
	case ColorNone:
		return "none"
	case ColorRed:
		return "红色"
	case ColorGreen:
		return "绿色"
	case ColorBlue:
		return "蓝色"
	case ColorYellow:
		return "黄色"
	default:
		return "none"
	}
}

func ParseColor(colorName string) int {
	switch colorName {
	case "none":
		return ColorNone
	case "红色":
		return ColorRed
	case "绿色":
		return ColorGreen
	case "蓝色":
		return ColorBlue
	case "黄色":
		return ColorYellow
	default:
		return ColorNone
	}
}
func hex2RGB(hex string) RGB {
	r, _ := strconv.ParseInt(hex[2:3], 16, 0)
	g, _ := strconv.ParseInt(hex[4:5], 16, 0)
	b, _ := strconv.ParseInt(hex[6:7], 16, 0)
	return RGB{uint8(r), uint8(g), uint8(b)}
}

func colorDistance(c1, c2 RGB) float64 {
	rDiff := float64(c1.R) - float64(c2.R)
	gDiff := float64(c1.G) - float64(c2.G)
	bDiff := float64(c1.B) - float64(c2.B)
	return math.Sqrt(rDiff*rDiff + gDiff*gDiff + bDiff*bDiff)
}

func findClosestColor(target RGB) RGB {
	var closest RGB
	minDistance := math.MaxFloat64
	for _, color := range colors {
		distance := colorDistance(target, color)
		if distance < minDistance {
			minDistance = distance
			closest = color
		}
	}
	return closest
}

func rgb2Color(color RGB) int {
	switch color {
	case Red:
		return ColorRed
	case Yellow:
		return ColorYellow
	case Blue:
		return ColorBlue
	case Green:
		return ColorGreen
	default:
		return ColorNone
	}
}

// ParseBgColor 解析表格颜色
func ParseBgColor(xlsx *XLSXFile) (map[int]int, error) {
	// 解析样式
	var stylesBytes []byte
	for _, f := range xlsx.Z.File {
		if f.Name == "xl/styles.xml" {
			rc, _ := f.Open()
			stylesBytes, _ = io.ReadAll(rc)
			rc.Close()
			break
		}
	}
	var s Styles
	if err := xml.Unmarshal(stylesBytes, &s); err != nil {
		return nil, err
	}
	colorsM := make(map[int]int)
	for i, dxf := range s.Dxfs.Dxfs {
		if dxf.Fill != nil && dxf.Fill.PatternFill != nil {
			fill := dxf.Fill.PatternFill
			if fill.BgColor.Rgb != "" {
				rgb := hex2RGB(fill.BgColor.Rgb)
				colorsM[i] = rgb2Color(findClosestColor(rgb))
			}
		}
	}
	return colorsM, nil
}
