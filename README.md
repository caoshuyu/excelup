通过配置读取/写入Excel
===

获取Excel表头
---
```go
    eUp := GetExcelUp()
	err := eUp.DecodeStyleJson(styleJson)
	if err != nil {
		fmt.Println(err)
		return
	}
	err = eUp.InitHeader()
	if err != nil {
		fmt.Println(err)
		return
	}
	//fileByte, err := eUp.GetFileBytes()
	//if err != nil {
	//	fmt.Println(err)
	//}
	err = eUp.SaveFile("./test.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
```

获取Excel数据
---
```go
data := &exceldata.ExcelFile{}
	sheetList := make([]*exceldata.ExcelSheet, 0)
	oneSheet := &exceldata.ExcelSheet{
		SheetName: "testSheet",
	}
	row1 := &exceldata.ExcelRow{}
	row2 := &exceldata.ExcelRow{}

	cellList := make([]*exceldata.ExcelCell, 0)
	cellList = append(cellList,
		&exceldata.ExcelCell{
			Key:   "class",
			Value: "一班",
		},
		&exceldata.ExcelCell{
			Key:   "name",
			Value: "王五",
		},
		&exceldata.ExcelCell{
			Key:   "age",
			Value: "10",
		},
		&exceldata.ExcelCell{
			Key:   "chinese",
			Value: "80",
		},
		&exceldata.ExcelCell{
			Key:   "mathematics",
			Value: "90",
		},
		&exceldata.ExcelCell{
			Key:   "english",
			Value: "80",
		},
		&exceldata.ExcelCell{
			Key:   "physical",
			Value: "99",
		},
		&exceldata.ExcelCell{
			Key:   "biological",
			Value: "100",
		},
		&exceldata.ExcelCell{
			Key:   "chemical",
			Value: "80",
		},
		&exceldata.ExcelCell{
			Key:   "geographic",
			Value: "90",
		},
		&exceldata.ExcelCell{
			Key:   "history",
			Value: "70",
		},
		&exceldata.ExcelCell{
			Key:   "political",
			Value: "60",
		},
	)

	row1.CellList = cellList
	row2.CellList = cellList
	oneSheet.RowList = append(oneSheet.RowList, row1, row2)
	sheetList = append(sheetList, oneSheet)
	data.SheetList = sheetList

	eUp := GetExcelUp()
	err := eUp.DecodeStyleJson(styleJson)
	if err != nil {
		fmt.Println(err)
		return
	}
	eUp.SheetData = data
	err = eUp.ExportExcel()
	if err != nil {
		fmt.Println(err)
		return
	}
	err = eUp.SaveFile("./test2.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
```


通过表头设置Excel数据
---
```go
	eUp := GetExcelUp()
	err := eUp.GetFile("./test2.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	err = eUp.DecodeStyleJson(styleJson)
	if err != nil {
		fmt.Println(err)
		return
	}
	err = eUp.ImportExcel()
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(eUp.SheetData)
```


配置文件说明
---
+ 可配置信息在excelstyle/style.go下
+ 时间类型需要在header中配置filed_type 为 date或者datetime，否则会被解析成小数

配置文件事例
---
```json
{
    "sheet_style":[
        {
            "sheet_name":"testSheet",
            "sheet_key":"",
            "sheet_header":{
                "header_line":3,
                "header_fields":[
                    {
                        "name":"班级",
                        "key":"class",
                        "filed_type":"string",
                        "merge_cell":{
                            "h_merge":0,
                            "v_merge":3
                        },
                        "cell_style":{
                            "border":{
                                "bottom":1,
                                "bottom_color":"FFFFFF00",
                                "top":1,
                                "top_color":"FFFFFF00",
                                "left":1,
                                "left_color":"FFFFFF00",
                                "right":1,
                                "right_color":"FFFFFF00"
                            }
                        },
                        "ad_same_value_merge":true
                    },
                    {
                        "name":"姓名",
                        "key":"name",
                        "filed_type":"string",
                        "merge_cell":{
                            "h_merge":0,
                            "v_merge":3
                        },
                        "cell_style":{
                            "border":{
                                "bottom":1,
                                "bottom_color":"FFFFFF00",
                                "top":1,
                                "top_color":"FFFFFF00",
                                "left":1,
                                "left_color":"FFFFFF00",
                                "right":1,
                                "right_color":"FFFFFF00"
                            }
                        }
                    },
                    {
                        "name":"年龄",
                        "key":"age",
                        "filed_type":"string",
                        "merge_cell":{
                            "h_merge":0,
                            "v_merge":3
                        },
                        "cell_style":{
                            "border":{
                                "bottom":1,
                                "bottom_color":"FFFFFF00",
                                "top":1,
                                "top_color":"FFFFFF00",
                                "left":1,
                                "left_color":"FFFFFF00",
                                "right":1,
                                "right_color":"FFFFFF00"
                            }
                        }
                    },
                    {
                        "name":"成绩",
                        "filed_type":"string",
                        "merge_cell":{
                            "v_merge":1
                        },
                        "cell_style":{
                            "border":{
                                "bottom":1,
                                "bottom_color":"FFFFFF00",
                                "top":1,
                                "top_color":"FFFFFF00",
                                "left":1,
                                "left_color":"FFFFFF00",
                                "right":1,
                                "right_color":"FFFFFF00"
                            }
                        },
                        "children":[
                            {
                                "name":"主科",
                                "cell_style":{
                                    "border":{
                                        "bottom":1,
                                        "bottom_color":"FFFFFF00",
                                        "top":1,
                                        "top_color":"FFFFFF00",
                                        "left":1,
                                        "left_color":"FFFFFF00",
                                        "right":1,
                                        "right_color":"FFFFFF00"
                                    }
                                },
                                "children":[
                                    {
                                        "name":"语文",
                                        "key":"chinese",
                                        "filed_type":"string",
                                        "cell_style":{
                                            "border":{
                                                "bottom":1,
                                                "bottom_color":"FFFFFF00",
                                                "top":1,
                                                "top_color":"FFFFFF00",
                                                "left":1,
                                                "left_color":"FFFFFF00",
                                                "right":1,
                                                "right_color":"FFFFFF00"
                                            },
                                            "font":{
                                                "color":"FFFF0000"
                                            }
                                        }
                                    },
                                    {
                                        "name":"数学",
                                        "key":"mathematics",
                                        "filed_type":"string",
                                        "cell_style":{
                                            "border":{
                                                "bottom":1,
                                                "bottom_color":"FFFFFF00",
                                                "top":1,
                                                "top_color":"FFFFFF00",
                                                "left":1,
                                                "left_color":"FFFFFF00",
                                                "right":1,
                                                "right_color":"FFFFFF00"
                                            },
                                            "font":{
                                                "color":"FF00B0F0"
                                            }
                                        }
                                    },
                                    {
                                        "name":"英语",
                                        "key":"english",
                                        "filed_type":"string",
                                        "cell_style":{
                                            "border":{
                                                "bottom":1,
                                                "bottom_color":"FFFFFF00",
                                                "top":1,
                                                "top_color":"FFFFFF00",
                                                "left":1,
                                                "left_color":"FFFFFF00",
                                                "right":1,
                                                "right_color":"FFFFFF00"
                                            }
                                        }
                                    }
                                ]
                            },
                            {
                                "name":"理科",
                                "cell_style":{
                                    "border":{
                                        "bottom":1,
                                        "bottom_color":"FFFFFF00",
                                        "top":1,
                                        "top_color":"FFFFFF00",
                                        "left":1,
                                        "left_color":"FFFFFF00",
                                        "right":1,
                                        "right_color":"FFFFFF00"
                                    }
                                },
                                "children":[
                                    {
                                        "name":"物理",
                                        "key":"physical",
                                        "filed_type":"string",
                                        "cell_style":{
                                            "border":{
                                                "bottom":1,
                                                "bottom_color":"FFFFFF00",
                                                "top":1,
                                                "top_color":"FFFFFF00",
                                                "left":1,
                                                "left_color":"FFFFFF00",
                                                "right":1,
                                                "right_color":"FFFFFF00"
                                            }
                                        }
                                    },
                                    {
                                        "name":"生物",
                                        "key":"biological",
                                        "filed_type":"string",
                                        "cell_style":{
                                            "border":{
                                                "bottom":1,
                                                "bottom_color":"FFFFFF00",
                                                "top":1,
                                                "top_color":"FFFFFF00",
                                                "left":1,
                                                "left_color":"FFFFFF00",
                                                "right":1,
                                                "right_color":"FFFFFF00"
                                            }
                                        }
                                    },
                                    {
                                        "name":"化学",
                                        "key":"chemical",
                                        "filed_type":"string",
                                        "cell_style":{
                                            "border":{
                                                "bottom":1,
                                                "bottom_color":"FFFFFF00",
                                                "top":1,
                                                "top_color":"FFFFFF00",
                                                "left":1,
                                                "left_color":"FFFFFF00",
                                                "right":1,
                                                "right_color":"FFFFFF00"
                                            }
                                        }
                                    }
                                ]
                            },
                            {
                                "name":"文科",
                                "cell_style":{
                                    "border":{
                                        "bottom":1,
                                        "bottom_color":"FFFFFF00",
                                        "top":1,
                                        "top_color":"FFFFFF00",
                                        "left":1,
                                        "left_color":"FFFFFF00",
                                        "right":1,
                                        "right_color":"FFFFFF00"
                                    }
                                },
                                "children":[
                                    {
                                        "name":"地理",
                                        "key":"geographic",
                                        "filed_type":"string",
                                        "cell_style":{
                                            "border":{
                                                "bottom":1,
                                                "bottom_color":"FFFFFF00",
                                                "top":1,
                                                "top_color":"FFFFFF00",
                                                "left":1,
                                                "left_color":"FFFFFF00",
                                                "right":1,
                                                "right_color":"FFFFFF00"
                                            }
                                        }
                                    },
                                    {
                                        "name":"历史",
                                        "key":"history",
                                        "filed_type":"string",
                                        "cell_style":{
                                            "border":{
                                                "bottom":1,
                                                "bottom_color":"FFFFFF00",
                                                "top":1,
                                                "top_color":"FFFFFF00",
                                                "left":1,
                                                "left_color":"FFFFFF00",
                                                "right":1,
                                                "right_color":"FFFFFF00"
                                            }
                                        }
                                    },
                                    {
                                        "name":"政治",
                                        "key":"political",
                                        "filed_type":"string",
                                        "cell_style":{
                                            "border":{
                                                "bottom":1,
                                                "bottom_color":"FFFFFF00",
                                                "top":1,
                                                "top_color":"FFFFFF00",
                                                "left":1,
                                                "left_color":"FFFFFF00",
                                                "right":1,
                                                "right_color":"FFFFFF00"
                                            }
                                        }
                                    }
                                ]
                            }
                        ]
                    }
                ]
            },
            "sheet_rows":[
                {
                    "row_num":10,
                    "weight":10,
                    "cell_style":{
                        "font":{
                            "color":"FF0000"
                        }
                    }
                },
                {
                    "row_num":-1,
                    "weight":10,
                    "cell_style":{
                        "font":{
                            "color":"CC00FF"
                        }
                    }
                }
            ],
            "sheet_cells":[
                {
                    "cell_num":-2,
                    "weight":12,
                    "cell_style":{
                        "font":{
                            "color":"66CCFF"
                        }
                    }
                }
            ],
            "sheet_row_cell":[
                {
                    "row_num":-2,
                    "cell_num":-1,
                    "cell_style":{
                        "font":{
                            "color":"6633FF"
                        }
                    }
                }
            ]
        }
    ]
}
```