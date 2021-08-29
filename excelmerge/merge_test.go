package excelmerge

import (
	"fmt"
	"testing"
)

func TestMergeHeader(t *testing.T) {
	styleJson := `{"sheet_style":[{"sheet_name":"testSheet","sheet_key":"","sheet_header":{"header_line":3,"header_fields":[{"name":"班级","key":"class","filed_type":"string","merge_cell":{"h_merge":0,"v_merge":3},"cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}},"ad_same_value_merge":true},{"name":"姓名","key":"name","filed_type":"string","merge_cell":{"h_merge":0,"v_merge":3},"cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}}},{"name":"年龄","key":"age","filed_type":"string","merge_cell":{"h_merge":0,"v_merge":3},"cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}}},{"name":"成绩","filed_type":"string","merge_cell":{"v_merge":1},"cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}},"children":[{"name":"主科","cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}},"children":[{"name":"语文","key":"chinese","filed_type":"string","cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"},"font":{"color":"FFFF0000"}}},{"name":"数学","key":"mathematics","filed_type":"string","cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"},"font":{"color":"FF00B0F0"}}},{"name":"英语","key":"english","filed_type":"string","cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}}}]},{"name":"理科","cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}},"children":[{"name":"物理","key":"physical","filed_type":"string","cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}}},{"name":"生物","key":"biological","filed_type":"string","cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}}},{"name":"化学","key":"chemical","filed_type":"string","cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}}}]},{"name":"文科","cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}},"children":[{"name":"地理","key":"geographic","filed_type":"string","cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}}},{"name":"历史","key":"history","filed_type":"string","cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}}},{"name":"政治","key":"political","filed_type":"string","cell_style":{"border":{"bottom":1,"bottom_color":"FFFFFF00","top":1,"top_color":"FFFFFF00","left":1,"left_color":"FFFFFF00","right":1,"right_color":"FFFFFF00"}}}]}]}]}}],"common_style":{}}`
	eUp := GetExcelUp()
	err := eUp.DecodeStyleJson(styleJson)
	if err != nil {
		fmt.Println(err)
	}
	err = eUp.MergeHeader()
	if err != nil {
		fmt.Println(err)
	}
	//fileByte, err := eUp.GetFileBytes()
	//if err != nil {
	//	fmt.Println(err)
	//}
	err = eUp.SaveFile("./test.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
