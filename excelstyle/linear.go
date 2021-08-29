package excelstyle

// 线型
const (
	LinearHair             = 1
	LinearDotted           = 2
	LinearDashDotDot       = 3
	LinearDashDot          = 4
	LinearDashed           = 5
	LinearThin             = 6
	LinearMediumDashDotDot = 7
	LinearSlantDashDot     = 8
	LinearMediumDashDot    = 9
	LinearMediumDashed     = 10
	LinearMedium           = 11
	LinearThick            = 12
	LinearDouble           = 13
)

// 获取线型名称
func GetLinearName(linearNum int) string {
	var linearName string
	switch linearNum {
	case LinearHair:
		linearName = "hair"
	case LinearDotted:
		linearName = "dotted"
	case LinearDashDotDot:
		linearName = "dashDotDot"
	case LinearDashDot:
		linearName = "dashDot"
	case LinearDashed:
		linearName = "dashed"
	case LinearThin:
		linearName = "thin"
	case LinearMediumDashDotDot:
		linearName = "mediumDashDotDot"
	case LinearSlantDashDot:
		linearName = "slantDashDot"
	case LinearMediumDashDot:
		linearName = "mediumDashDot"
	case LinearMediumDashed:
		linearName = "mediumDashed"
	case LinearMedium:
		linearName = "medium"
	case LinearThick:
		linearName = "thick"
	case LinearDouble:
		linearName = "double"
	}
	return linearName
}
