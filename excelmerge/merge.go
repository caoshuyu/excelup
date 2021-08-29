package excelmerge

import (
	"bytes"
	"encoding/json"
	"errors"
	"github.com/caoshuyu/excelup/exceldata"
	"github.com/caoshuyu/excelup/excelstyle"
	"github.com/tealeg/xlsx"
	"strconv"
)

// Excel 结构体
type ExcelUp struct {
	File          *xlsx.File
	Style         *excelstyle.ExcelStyle
	HeardKeyMap   map[string]*HeardKeyInfo
	HeardNameMap  map[string]*HeardKeyInfo
	HeardIndexMap map[int]*HeardKeyInfo
	SheetData     []*exceldata.ExcelFile
}

type HeardKeyInfo struct {
	Name             string // 字段名
	Key              string // 字段key
	Index            int    // 字段所在cell位置
	FiledType        string // 字段类型
	AdSameValueMerge bool   // 类似字段合并
}

func GetExcelUp() *ExcelUp {
	eUp := &ExcelUp{}
	eUp.File = xlsx.NewFile()
	eUp.Style = new(excelstyle.ExcelStyle)
	eUp.HeardKeyMap = make(map[string]*HeardKeyInfo)
	eUp.HeardNameMap = make(map[string]*HeardKeyInfo)
	eUp.HeardIndexMap = make(map[int]*HeardKeyInfo)

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

// 合并头信息
func (e *ExcelUp) MergeHeader() (err error) {
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
				cellIndex = e._getChildKey(sheetStyle.SheetHeader.HeaderFields[k].Children, cellIndex)
			} else {
				e._setHeardKeyInfo(sheetStyle.SheetHeader.HeaderFields[k], cellIndex)
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

func (e *ExcelUp) _getChildKey(children []*excelstyle.SheetHeaderField, cellIndex int) (newCellIndex int) {
	for childK := range children {
		if len(children[childK].Children) > 0 {
			cellIndex = e._getChildKey(children[childK].Children, cellIndex)
		} else {
			e._setHeardKeyInfo(children[childK], cellIndex)
			cellIndex++
		}
	}
	return cellIndex
}

func (e *ExcelUp) _setHeardKeyInfo(headerFields *excelstyle.SheetHeaderField, cellIndex int) {
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

	e.HeardKeyMap[hkInfo.Key] = hkInfo
	e.HeardNameMap[hkInfo.Name] = hkInfo
	e.HeardIndexMap[hkInfo.Index] = hkInfo
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
