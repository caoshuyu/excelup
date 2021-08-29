package excelstyle

// Excel 样式信息
type ExcelStyle struct {
	SheetStyle  []*ExcelSheetStyle `json:"sheet_style"`  // sheet设置
	CommonStyle *ExcelCommonStyle  `json:"common_style"` // 通用sheet设置
}

// Excel sheet 样式信息
type ExcelSheetStyle struct {
	SheetName    string            `json:"sheet_name"`     // sheet名称
	SheetKey     string            `json:"sheet_key"`      // 数据映射名称
	SheetHeader  *SheetHeader      `json:"sheet_header"`   // 表头设置
	SheetRows    []*RowStyle       `json:"sheet_rows"`     // 行设置
	SheetCells   []*CellStyle      `json:"sheet_cells"`    // 列设置
	SheetRowCell []*RowCellStyle   `json:"sheet_row_cell"` // 行列设置
	MergeCell    []*MergeCellStyle `json:"merge_cell"`     // 合并设置
}

// cell 合并信息
type MergeCellStyle struct {
	RowNum  int `json:"row_num"`  // 行号
	CellNum int `json:"cell_num"` // 列号
	HMerge  int `json:"h_merge"`  // 列合并
	VMerge  int `json:"v_merge"`  // 行合并
}

// Excel 通用样式信息
type ExcelCommonStyle struct {
	AllRows  []*RowStyle  `json:"all_rows"`  // 行配置
	AllCells []*CellStyle `json:"all_cells"` // 列配置
}

type SheetHeader struct {
	HeaderLine   int                 `json:"header_line"`   // 表头所在行
	HeaderFields []*SheetHeaderField `json:"header_fields"` // 表头字段
	RemoveHeader bool                `json:"remove_header"` // 导入数据时删除sheet表头行
	AddHeader    bool                `json:"add_header"`    // 导出数据时添加header表头行
}

type SheetHeaderField struct {
	Name             string              `json:"name"`                // 字段名
	Key              string              `json:"key"`                 // 数据时映射名称
	FiledType        string              `json:"filed_type"`          // 字段类型
	MergeCell        *MergeCellStyle     `json:"merge_cell"`          // 合并信息
	CellStyle        *CellStyleValue     `json:"cell_style"`          // 单元格样式
	Children         []*SheetHeaderField `json:"children"`            // 子字段，主字段列数大于等于子字段列数和
	AdSameValueMerge bool                `json:"ad_same_value_merge"` // 临近相同值合并
}

type RowStyle struct {
	RowNum int     `json:"row_num"` // 行数，-1奇数行，-2偶数行
	Weight int     `json:"weight"`  // 权重，row，cell交叉时权重高的覆盖权重低的，同级配置内生效
	Height float64 `json:"height"`  // 高度
}

type CellStyle struct {
	CellNum   int             `json:"cell_num"`   // 行数，-1奇数行，-2偶数行
	Weight    int             `json:"weight"`     // 权重，row，cell交叉时权重高的覆盖权重低的，同级配置内生效
	Width     float64         `json:"width"`      // 宽度
	CellStyle *CellStyleValue `json:"cell_style"` // 单元格样式
}

type RowCellStyle struct {
	RowNum    int             `json:"row_num"`    // 行数，-1奇数行，-2偶数行
	CellNum   int             `json:"cell_num"`   // 行数，-1奇数行，-2偶数行
	CellStyle *CellStyleValue `json:"cell_style"` // 单元格样式
}

// 字段值
type CellStyleValue struct {
	Border    *StyleBorder    `json:"border"`    // 边框
	Font      *StyleFont      `json:"font"`      // 字体
	Fill      *StyleFill      `json:"fill"`      // 背景
	Alignment *StyleAlignment `json:"alignment"` // 文字样式
}

type StyleBorder struct {
	Bottom      int    `json:"bottom"`       // 下边框样式
	BottomColor string `json:"bottom_color"` // 下边框颜色
	Top         int    `json:"top"`          // 上边框样式
	TopColor    string `json:"top_color"`    // 上边框颜色
	Left        int    `json:"left"`         // 左边框样式
	LeftColor   string `json:"left_color"`   // 左边框颜色
	Right       int    `json:"right"`        // 右边框样式
	RightColor  string `json:"right_color"`  // 右边框颜色
}

type StyleFont struct {
	Name      string `json:"name"`      // 字体名称
	Bold      bool   `json:"bold"`      // 是否加粗
	Charset   int    `json:"charset"`   // 字符集
	Color     string `json:"color"`     // 字体颜色
	Family    int    `json:"family"`    //
	Italic    bool   `json:"italic"`    // 斜体
	Size      int    `json:"size"`      // 字号
	Underline bool   `json:"underline"` // 下划线
}

type StyleFill struct {
	BgColor     string `json:"bg_color"`     // 背景色
	FgColor     string `json:"fg_color"`     // 图案颜色
	PatternType int    `json:"pattern_type"` // 图案样式
}

type StyleAlignment struct {
	Horizontal   string `json:"horizontal"`    // 水平对其方式
	Indent       int    `json:"indent"`        // 缩进，读取数据时无法读取
	ShrinkToFit  bool   `json:"shrink_to_fit"` // 缩小以适应
	TextRotation int    `json:"text_rotation"` // 文本旋转
	Vertical     string `json:"vertical"`      // 垂直对其方式
	WrapText     bool   `json:"wrap_text"`     // 文字换行
}
