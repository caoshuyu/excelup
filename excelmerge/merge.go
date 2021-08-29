package excelmerge

import (
	"bytes"
	"encoding/json"
	"errors"
	"github.com/caoshuyu/excelup/exceldata"
	"github.com/caoshuyu/excelup/excelstyle"
	"github.com/tealeg/xlsx"
	"path"
	"strconv"
	"time"
)

// Excel 结构体
type ExcelUp struct {
	File                 *xlsx.File
	FileName             string
	Style                *excelstyle.ExcelStyle
	HeardKeyMap          map[string]map[string]*HeardKeyInfo
	HeardNameMap         map[string]map[string]*HeardKeyInfo
	HeardIndexMap        map[string]map[int]*HeardKeyInfo
	HeardHigh            map[string]int
	DataStartLine        map[string]int
	SheetData            *exceldata.ExcelFile
	SheetRowStyleMap     map[string]map[int]*excelstyle.RowStyle
	SheetCellStyleMap    map[string]map[int]*excelstyle.CellStyle
	SheetRowCellStyleMap map[string]map[string]*excelstyle.RowCellStyle
}

type HeardKeyInfo struct {
	Name             string // 字段名
	Key              string // 字段key
	Index            int    // 字段所在cell位置,从1开始
	FiledType        string // 字段类型
	AdSameValueMerge bool   // 类似字段合并
}

func GetExcelUp() *ExcelUp {
	eUp := &ExcelUp{}
	eUp.File = xlsx.NewFile()
	eUp.Style = new(excelstyle.ExcelStyle)
	eUp.HeardKeyMap = make(map[string]map[string]*HeardKeyInfo)
	eUp.HeardNameMap = make(map[string]map[string]*HeardKeyInfo)
	eUp.HeardIndexMap = make(map[string]map[int]*HeardKeyInfo)
	eUp.HeardHigh = make(map[string]int)
	eUp.DataStartLine = make(map[string]int)
	eUp.SheetRowStyleMap = make(map[string]map[int]*excelstyle.RowStyle)
	eUp.SheetCellStyleMap = make(map[string]map[int]*excelstyle.CellStyle)
	eUp.SheetRowCellStyleMap = make(map[string]map[string]*excelstyle.RowCellStyle)

	return eUp
}

func (e *ExcelUp) DecodeStyleJson(styleJson string) error {
	style := &excelstyle.ExcelStyle{}
	err := json.Unmarshal([]byte(styleJson), style)
	if err != nil {
		return err
	}
	e.Style = style
	return nil
}

// 获取文件流
func (e *ExcelUp) GetFileBytes() (fileByte []byte, err error) {
	buffer := new(bytes.Buffer)
	err = e.File.Write(buffer)
	if err != nil {
		return nil, err
	}
	return buffer.Bytes(), nil
}

// 存储文件
func (e *ExcelUp) SaveFile(filePath string) (err error) {
	return e.File.Save(filePath)
}

// 读取文件
func (e *ExcelUp) GetFile(filePath string) (err error) {
	e.FileName = path.Base(filePath)
	file, err := xlsx.OpenFile(filePath)
	if nil != err {
		return
	}
	e.File = file
	return nil
}

// 初始化头信息
func (e *ExcelUp) InitHeader() (err error) {
	// 计算头信息
	if e.File == nil {
		return errors.New("excel file nil")
	}
	if e.Style == nil || e.Style.SheetStyle == nil {
		return
	}

	for _, sheetStyle := range e.Style.SheetStyle {
		if sheetStyle.SheetHeader == nil {
			continue
		}
		// 计算头部占用行列
		maxRowNum, maxCellNum := e._calculateHeaderRowCellNum(sheetStyle)
		if maxRowNum == 0 || maxCellNum == 0 {
			continue
		}
		e.HeardHigh[sheetStyle.SheetName] = maxRowNum
		// 生成头部信息
		fileSheet, h := e.File.Sheet[sheetStyle.SheetName]
		if !h {
			fileSheet, err = e.File.AddSheet(sheetStyle.SheetName)
			if err != nil {
				return err
			}
		}
		// 头部空行
		emptyRow := 0
		if sheetStyle.SheetHeader.HeaderLine > 1 {
			emptyRow = sheetStyle.SheetHeader.HeaderLine - 1
			for i := 0; i < emptyRow; i++ {
				fileSheet.AddRow()
			}
		}
		e.DataStartLine[sheetStyle.SheetName] = maxRowNum + emptyRow

		// 构建数据存储集合
		for i := 0; i < maxRowNum; i++ {
			row := fileSheet.AddRow()
			for j := 0; j < maxCellNum; j++ {
				row.AddCell()
			}
		}
		// 存储数据和样式
		cellIndex := 1
		for k := range sheetStyle.SheetHeader.HeaderFields {
			if len(sheetStyle.SheetHeader.HeaderFields[k].Children) > 0 {
				cellIndex = e._getChildKey(sheetStyle.SheetName, sheetStyle.SheetHeader.HeaderFields[k].Children, cellIndex)
			} else {
				e._setHeardKeyInfo(sheetStyle.SheetName, sheetStyle.SheetHeader.HeaderFields[k], cellIndex)
				cellIndex++
			}
		}
		// 填充头部信息
		cellMap := make(map[string]*xlsx.Cell)
		for rowIndex := range fileSheet.Rows {
			for cellIndex := range fileSheet.Rows[rowIndex].Cells {
				cellMap[strconv.Itoa(rowIndex)+"_"+strconv.Itoa(cellIndex)] = fileSheet.Rows[rowIndex].Cells[cellIndex]
			}
		}

		nowRow := sheetStyle.SheetHeader.HeaderLine - 1
		nowCell := 0
		for _, headerField := range sheetStyle.SheetHeader.HeaderFields {
			startCell := nowCell
			if headerField.MergeCell == nil {
				headerField.MergeCell = &excelstyle.MergeCellStyle{}
			}
			cellNow := cellMap[strconv.Itoa(nowRow)+"_"+strconv.Itoa(nowCell)]
			if cellNow == nil {
				continue
			}
			cellNow.Value = headerField.Name
			cellNow.VMerge = headerField.MergeCell.VMerge
			cellNow.HMerge = headerField.MergeCell.HMerge
			if len(headerField.Children) > 0 {
				moveCell := e._setChildHeard(nowRow+headerField.MergeCell.VMerge+1, nowCell, cellMap, headerField.Children)
				nowCell += moveCell
				if headerField.MergeCell.HMerge < moveCell-1 {
					cellNow.HMerge = moveCell - 1
				}
			} else {
				nowCell += headerField.MergeCell.HMerge + 1
			}
			// 行
			for rNo, endRNo := nowRow, nowRow+headerField.MergeCell.VMerge+1; rNo < endRNo; rNo++ {
				// 列
				for cNo := startCell; cNo < startCell+cellNow.HMerge+1; cNo++ {
					cell := cellMap[strconv.Itoa(rNo)+"_"+strconv.Itoa(cNo)]
					if cell == nil {
						continue
					}
					// 设置样式
					e._setRowCellStyle(cell, headerField.CellStyle)
				}
			}
		}
	}

	return nil
}

// 生成Excel文件
func (e *ExcelUp) ExportExcel() (err error) {
	// 生成sheet文件头
	if err = e.InitHeader(); err != nil {
		return err
	}
	if e.SheetData == nil {
		return
	}
	// 初始数据样式化配置信息
	e._initDataStyle()

	// 写入数据
	for _, sheet := range e.SheetData.SheetList {
		sheetName := sheet.SheetName
		oneSheet := e.File.Sheet[sheetName]
		if oneSheet == nil {
			oneSheet, err = e.File.AddSheet(sheetName)
			if err != nil {
				return err
			}
		}
		sheetKeyMap := e.HeardKeyMap[sheetName]

		for rowNum, rowData := range sheet.RowList {
			useRowNum := e.DataStartLine[sheetName] + rowNum + 1
			oneRow := oneSheet.AddRow()
			maxCellNum := 0
			for cellI, cellVal := range rowData.CellList {
				if len(cellVal.Key) > 0 {
					if hkInfo, h := sheetKeyMap[cellVal.Key]; h {
						if hkInfo.Index > maxCellNum {
							maxCellNum = hkInfo.Index
						}
					} else {
						if cellI+1 > maxCellNum {
							maxCellNum = cellI + 1
						}
					}
				}
			}
			// 创建cell map
			cellMap := make(map[int]*xlsx.Cell)
			for i := 1; i <= maxCellNum; i++ {
				cellMap[i] = oneRow.AddCell()
			}
			for cellNum, cellData := range rowData.CellList {
				useCellNum := cellNum + 1
				if len(cellData.Key) > 0 {
					if sheetKey, h := sheetKeyMap[cellData.Key]; h {
						useCellNum = sheetKey.Index
					}
				}
				cell := cellMap[useCellNum]
				if cell == nil {
					// 未在key初始化字段加到行尾
					cell = oneRow.AddCell()
				}
				cell.SetValue(cellData.Value)
				// 获取样式
				style := e._getDataStyle(sheetName, useRowNum, useCellNum)
				if style != nil {
					e._setRowCellStyle(cell, style)
				}
			}
		}
	}
	return nil
}

// 导入Excel
func (e *ExcelUp) ImportExcel() (err error) {
	// 生成sheet文件头
	if err = e.InitHeader(); err != nil {
		return err
	}
	// 解析数据
	sheetList := make([]*exceldata.ExcelSheet, 0)
	for _, sheet := range e.File.Sheets {
		sheetName := sheet.Name
		startLine := e.DataStartLine[sheetName] // 下标
		sheetIndexMap := e.HeardIndexMap[sheetName]

		oneSheet := &exceldata.ExcelSheet{
			SheetName: sheet.Name,
		}
		for index, row := range sheet.Rows {
			if index < startLine {
				continue
			}
			// 判别是否是空行
			isEmptyRow := true
			for _, cell := range row.Cells {
				if len(cell.Value) > 0 {
					isEmptyRow = false
					break
				}
			}
			if isEmptyRow {
				continue
			}
			oneRow := &exceldata.ExcelRow{}
			for cellIndex, cell := range row.Cells {
				hkInfo, h := sheetIndexMap[cellIndex+1]
				val := e._checkDateValue(cell.String(), cell.NumFmt, hkInfo.FiledType)
				oneCel := &exceldata.ExcelCell{
					Value: val,
				}
				if h {
					oneCel.Key = hkInfo.Key
				} else {
					oneCel.Key = strconv.Itoa(cellIndex + 1)
				}
				oneRow.CellList = append(oneRow.CellList, oneCel)
			}
			if oneRow.CellList == nil {
				continue
			}
			oneSheet.RowList = append(oneSheet.RowList, oneRow)
		}
		sheetList = append(sheetList, oneSheet)
	}
	e.SheetData = &exceldata.ExcelFile{
		FileName:  e.FileName,
		SheetList: sheetList,
	}
	return nil
}

func (e *ExcelUp) _initDataStyle() {
	for _, sheetStyle := range e.Style.SheetStyle {
		sheetName := sheetStyle.SheetName
		oneSheetRowStyleMap, h := e.SheetRowStyleMap[sheetName]
		if !h {
			oneSheetRowStyleMap = make(map[int]*excelstyle.RowStyle)
		}
		for k, sheetRow := range sheetStyle.SheetRows {
			oneSheetRowStyleMap[sheetRow.RowNum] = sheetStyle.SheetRows[k]
		}
		e.SheetRowStyleMap[sheetName] = oneSheetRowStyleMap

		oneSheetCellStyleMap, h := e.SheetCellStyleMap[sheetName]
		if !h {
			oneSheetCellStyleMap = make(map[int]*excelstyle.CellStyle)
		}
		for k, sheetCell := range sheetStyle.SheetCells {
			oneSheetCellStyleMap[sheetCell.CellNum] = sheetStyle.SheetCells[k]
		}
		e.SheetCellStyleMap[sheetName] = oneSheetCellStyleMap

		oneSheetRowCellStyleMap, h := e.SheetRowCellStyleMap[sheetName]
		if !h {
			oneSheetRowCellStyleMap = make(map[string]*excelstyle.RowCellStyle)
		}
		for k, rowCell := range sheetStyle.SheetRowCell {
			key := strconv.Itoa(rowCell.RowNum) + "_" + strconv.Itoa(rowCell.CellNum)
			oneSheetRowCellStyleMap[key] = sheetStyle.SheetRowCell[k]
		}
		e.SheetRowCellStyleMap[sheetName] = oneSheetRowCellStyleMap
	}
}

// 检测数据类型，时间类型转换为YYYY-MM-DD HH:mm:ss
func (e *ExcelUp) _checkDateValue(data string, numFmt string, fileType string) string {
	val := ""
	if fileType == "date" || fileType == "datetime" {
		i, e := strconv.ParseFloat(data, 64)
		if e != nil {
			val = data
		}
		day := int(i)
		compensationDay := 1
		if day > 60 {
			// excel bug 1900年有2-29日，此日期不存在
			compensationDay = 2
		}
		second := int((i-float64(day))*86400 + 0.5)
		tm := time.Date(1900, 1, 1, 0, 0, 0, 0,
			time.FixedZone("Asia/Shanghai", 0))
		tm = tm.Add(time.Hour * 24 * time.Duration(day-compensationDay))
		tm = tm.Add(time.Second * time.Duration(second))
		if fileType == "date" {
			val = tm.Format("2006-01-02")
		} else {
			val = tm.Format("2006-01-02 15:04:05")
		}
	} else {
		val = data
	}
	return val
}

// 获取数据样式
func (e *ExcelUp) _getDataStyle(sheetName string, rowNum, cellNum int) *excelstyle.CellStyleValue {
	rowParity := -1  // 行奇偶
	cellParity := -1 // 列奇偶
	if rowNum%2 == 0 {
		rowParity = -2
	}
	if cellNum%2 == 0 {
		cellParity = -2
	}
	// 获取 row cell 样式
	if sheetStyleMap, h := e.SheetRowCellStyleMap[sheetName]; h {
		// row ,cell 值
		rowCellKey := strconv.Itoa(rowNum) + "_" + strconv.Itoa(cellNum)
		if style, h := sheetStyleMap[rowCellKey]; h {
			return style.CellStyle
		}
		// row 值 , cell 范围
		rowCellKey = strconv.Itoa(rowNum) + "_" + strconv.Itoa(cellParity)
		if style, h := sheetStyleMap[rowCellKey]; h {
			return style.CellStyle
		}
		// row 范围 , cell 值
		rowCellKey = strconv.Itoa(rowParity) + "_" + strconv.Itoa(cellNum)
		if style, h := sheetStyleMap[rowCellKey]; h {
			return style.CellStyle
		}
		// row 范围 , cell 范围
		rowCellKey = strconv.Itoa(rowParity) + "_" + strconv.Itoa(cellParity)
		if style, h := sheetStyleMap[rowCellKey]; h {
			return style.CellStyle
		}
	}

	var styleValue *excelstyle.CellStyleValue
	var wight int
	// row / cell 获取数据比权重决定
	// 获取 row 样式
	if sheetStyleMap, h := e.SheetRowStyleMap[sheetName]; h {
		// row 值
		if style, h := sheetStyleMap[rowNum]; h {
			if style.Weight > wight {
				styleValue = style.CellStyle
				wight = style.Weight
			}
		} else {
			// row 范围
			if style, h = sheetStyleMap[rowParity]; h {
				if style.Weight > wight {
					styleValue = style.CellStyle
					wight = style.Weight
				}
			}
		}
	}
	// 获取 cell 样式
	if sheetStyleMap, h := e.SheetCellStyleMap[sheetName]; h {
		if style, h := sheetStyleMap[cellNum]; h {
			// cell 值
			if style.Weight > wight {
				styleValue = style.CellStyle
				wight = style.Weight
			}
		} else {
			// cell 范围
			if style, h := sheetStyleMap[cellParity]; h {
				if style.Weight > wight {
					styleValue = style.CellStyle
					wight = style.Weight
				}
			}
		}
	}
	return styleValue
}

func (e *ExcelUp) _setChildHeard(nowRow, nowCell int, cellMap map[string]*xlsx.Cell, children []*excelstyle.SheetHeaderField) int {
	moveCell := 0
	useCellNo := nowCell
	for _, child := range children {
		// 处理当前节点信息
		cellNow := cellMap[strconv.Itoa(nowRow)+"_"+strconv.Itoa(useCellNo)]
		if cellNow == nil {
			continue
		}
		if child.MergeCell == nil {
			child.MergeCell = &excelstyle.MergeCellStyle{}
		}
		cellNow.Value = child.Name
		cellNow.VMerge = child.MergeCell.VMerge
		cellNow.HMerge = child.MergeCell.HMerge

		childMove := 0
		if len(child.Children) > 0 {
			childMove = e._setChildHeard(nowRow+child.MergeCell.VMerge+1, useCellNo, cellMap, child.Children)
		}
		if childMove-1 > child.MergeCell.HMerge {
			cellNow.HMerge = childMove - 1
		}
		if childMove > child.MergeCell.HMerge+1 {
			moveCell += childMove
		} else {
			moveCell += child.MergeCell.HMerge + 1
		}
		useCellNo = nowCell + moveCell
		// 设置样式
		for cRow := nowRow; cRow < nowRow+child.MergeCell.VMerge+1; cRow++ {
			for cCell := nowCell; cCell < nowCell+moveCell; cCell++ {
				styleCell := cellMap[strconv.Itoa(cRow)+"_"+strconv.Itoa(cCell)]
				e._setRowCellStyle(styleCell, child.CellStyle)
			}
		}
	}
	return moveCell
}

func (e *ExcelUp) _getChildKey(sheetName string, children []*excelstyle.SheetHeaderField, cellIndex int) (newCellIndex int) {
	for childK := range children {
		if len(children[childK].Children) > 0 {
			cellIndex = e._getChildKey(sheetName, children[childK].Children, cellIndex)
		} else {
			e._setHeardKeyInfo(sheetName, children[childK], cellIndex)
			cellIndex++
		}
	}
	return cellIndex
}

func (e *ExcelUp) _setHeardKeyInfo(sheetName string, headerFields *excelstyle.SheetHeaderField, cellIndex int) {
	hkInfo := &HeardKeyInfo{}
	if len(headerFields.Key) > 0 && len(headerFields.Name) > 0 {
		hkInfo.Name = headerFields.Name
		hkInfo.Key = headerFields.Key
		hkInfo.Index = cellIndex
	} else if len(headerFields.Key) > 0 {
		hkInfo.Name = headerFields.Key
		hkInfo.Key = headerFields.Key
		hkInfo.Index = cellIndex
	} else if len(headerFields.Name) > 0 {
		hkInfo.Name = headerFields.Name
		hkInfo.Key = headerFields.Name
		hkInfo.Index = cellIndex
	} else {
		hkInfo.Name = strconv.Itoa(cellIndex + 1)
		hkInfo.Key = strconv.Itoa(cellIndex + 1)
		hkInfo.Index = cellIndex
	}
	if len(headerFields.FiledType) > 0 {
		hkInfo.FiledType = headerFields.FiledType
	} else {
		hkInfo.FiledType = "nil"
	}
	hkInfo.AdSameValueMerge = headerFields.AdSameValueMerge
	keyMap, h := e.HeardKeyMap[sheetName]
	if !h {
		keyMap = make(map[string]*HeardKeyInfo)
	}
	keyMap[hkInfo.Key] = hkInfo
	e.HeardKeyMap[sheetName] = keyMap
	NameMap, h := e.HeardNameMap[sheetName]
	if !h {
		NameMap = make(map[string]*HeardKeyInfo)
	}
	NameMap[hkInfo.Name] = hkInfo
	e.HeardNameMap[sheetName] = NameMap
	indexMap, h := e.HeardIndexMap[sheetName]
	if !h {
		indexMap = make(map[int]*HeardKeyInfo)
	}
	indexMap[hkInfo.Index] = hkInfo
	e.HeardIndexMap[sheetName] = indexMap
}

func (e *ExcelUp) _calculateHeaderRowCellNum(sheetStyle *excelstyle.ExcelSheetStyle) (maxRowNum, maxCellNum int) {
	for _, headerFields := range sheetStyle.SheetHeader.HeaderFields {
		childRow, childCell := 0, 0
		if len(headerFields.Children) > 0 {
			childRow, childCell = e._calculateHeaderRowCellNumChildren(headerFields, maxRowNum, maxCellNum)
		}
		if headerFields.MergeCell != nil {
			if headerFields.MergeCell.VMerge+1+childRow > maxRowNum {
				maxRowNum = headerFields.MergeCell.VMerge + 1 + childRow
			}
			if headerFields.MergeCell.HMerge+1 > childCell {
				maxCellNum += headerFields.MergeCell.HMerge + 1
			} else {
				maxCellNum += childCell
			}
		} else {
			if maxRowNum == 0 {
				maxRowNum = 1
			}
			if maxRowNum < 1+childRow {
				maxRowNum = 1 + childRow
			}
			if childCell > 0 {
				maxCellNum += childCell
			} else {
				maxCellNum += 1
			}
		}
	}
	return maxRowNum, maxCellNum
}

func (e *ExcelUp) _calculateHeaderRowCellNumChildren(headerFields *excelstyle.SheetHeaderField, maxRowNum, maxCellNum int) (int, int) {
	childRow := 0
	childCell := 0
	for _, child := range headerFields.Children {
		if len(child.Children) > 0 {
			newRow, newCell := e._calculateHeaderRowCellNumChildren(child, maxRowNum, maxCellNum)
			if newRow > childRow {
				childRow = newRow
			}
			childCell += newCell
		} else {
			if child.MergeCell == nil {
				childCell += 1
				if childRow == 0 {
					childRow = 1
				}
			} else {
				childCell += child.MergeCell.HMerge + 1
				if child.MergeCell.VMerge+1 > childRow {
					childRow = child.MergeCell.VMerge + 1
				}
			}
		}
	}
	return childRow, childCell
}

func (e *ExcelUp) _setRowCellStyle(cell *xlsx.Cell, cellStyle *excelstyle.CellStyleValue) {
	style := new(xlsx.Style)
	if cellStyle.Border != nil {
		if cellStyle.Border.Bottom != 0 {
			style.Border.Bottom = excelstyle.GetLinearName(cellStyle.Border.Bottom)
		}
		if len(cellStyle.Border.BottomColor) > 0 {
			style.Border.BottomColor = cellStyle.Border.BottomColor
		}

		if cellStyle.Border.Top != 0 {
			style.Border.Top = excelstyle.GetLinearName(cellStyle.Border.Top)
		}
		if len(cellStyle.Border.TopColor) > 0 {
			style.Border.TopColor = cellStyle.Border.TopColor
		}

		if cellStyle.Border.Left != 0 {
			style.Border.Left = excelstyle.GetLinearName(cellStyle.Border.Left)
		}
		if len(cellStyle.Border.LeftColor) > 0 {
			style.Border.LeftColor = cellStyle.Border.LeftColor
		}

		if cellStyle.Border.Right != 0 {
			style.Border.Right = excelstyle.GetLinearName(cellStyle.Border.Right)
		}
		if len(cellStyle.Border.RightColor) > 0 {
			style.Border.RightColor = cellStyle.Border.RightColor
		}
	}
	if cellStyle.Font != nil {
		style.Font.Name = cellStyle.Font.Name
		style.Font.Bold = cellStyle.Font.Bold
		style.Font.Charset = cellStyle.Font.Charset
		style.Font.Color = cellStyle.Font.Color
		style.Font.Family = cellStyle.Font.Family
		style.Font.Italic = cellStyle.Font.Italic
		style.Font.Size = cellStyle.Font.Size
		style.Font.Underline = cellStyle.Font.Underline
	}
	if cellStyle.Fill != nil {
		if cellStyle.Fill.PatternType > 0 {
			style.Fill.PatternType = excelstyle.GetPatternName(cellStyle.Fill.PatternType)
		}
		if len(cellStyle.Fill.FgColor) > 0 {
			style.Fill.FgColor = cellStyle.Fill.FgColor
		}
		if len(cellStyle.Fill.BgColor) > 0 {
			style.Fill.BgColor = cellStyle.Fill.BgColor
		}
	}
	if cellStyle.Alignment != nil {
		if len(cellStyle.Alignment.Horizontal) > 0 {
			style.Alignment.Horizontal = cellStyle.Alignment.Horizontal
		}
		if len(cellStyle.Alignment.Vertical) > 0 {
			style.Alignment.Vertical = cellStyle.Alignment.Vertical
		}
		style.Alignment.Indent = cellStyle.Alignment.Indent
		style.Alignment.ShrinkToFit = cellStyle.Alignment.ShrinkToFit
		style.Alignment.TextRotation = cellStyle.Alignment.TextRotation
		style.Alignment.WrapText = cellStyle.Alignment.WrapText
	}
	cell.SetStyle(style)
}
