package exceldata

// excel 文件
type ExcelFile struct {
	FileName  string        `json:"file_name"`
	SheetList []*ExcelSheet `json:"sheet_list"`
}

// excel Sheet
type ExcelSheet struct {
	SheetName string      `json:"sheet_name"`
	RowList   []*ExcelRow `json:"row_list"`
}

// excel Row
type ExcelRow struct {
	CellList []*ExcelCell `json:"cell_list"`
}

// excel cell
type ExcelCell struct {
	Key   string `json:"key"`
	Value string `json:"value"`
}
