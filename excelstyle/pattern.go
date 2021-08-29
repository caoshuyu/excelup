package excelstyle

// 图案
const (
	PatternSolid           = 1
	PatternDarkGray        = 2
	PatternMediumGray      = 3
	PatternLightGray       = 4
	PatternGray125         = 5
	PatternGray0625        = 6
	PatternDarkHorizontal  = 7
	PatternDarkVertical    = 8
	PatternDarkDown        = 9
	PatternDarkUp          = 10
	PatternDarkGrid        = 11
	PatternDarkTrellis     = 12
	PatternLightHorizontal = 13
	PatternLightVertical   = 14
	PatternLightDown       = 15
	PatternLightUp         = 16
	PatternLightGrid       = 17
	PatternLightTrellis    = 18
)

// 获取图案名称
func GetPatternName(patternNum int) string {
	var patternName string
	switch patternNum {
	case PatternSolid:
		patternName = "solid"
	case PatternDarkGray:
		patternName = "darkGray"
	case PatternMediumGray:
		patternName = "mediumGray"
	case PatternLightGray:
		patternName = "lightGray"
	case PatternGray125:
		patternName = "gray125"
	case PatternGray0625:
		patternName = "gray0625"
	case PatternDarkHorizontal:
		patternName = "darkHorizontal"
	case PatternDarkVertical:
		patternName = "darkVertical"
	case PatternDarkDown:
		patternName = "darkDown"
	case PatternDarkUp:
		patternName = "darkUp"
	case PatternDarkGrid:
		patternName = "darkGrid"
	case PatternDarkTrellis:
		patternName = "darkTrellis"
	case PatternLightHorizontal:
		patternName = "lightHorizontal"
	case PatternLightVertical:
		patternName = "lightVertical"
	case PatternLightDown:
		patternName = "lightDown"
	case PatternLightUp:
		patternName = "lightUp"
	case PatternLightGrid:
		patternName = "lightGrid"
	case PatternLightTrellis:
		patternName = "lightTrellis"
	}
	return patternName
}
