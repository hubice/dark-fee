package main

import (
	"errors"
	"fmt"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

type Rule struct {
	Name string
	PeopleContent string
}

type Fee struct {
	f    *excelize.File
	isDebug bool
	isAll []string

	rows [][]string
	modeRows map[string][]Rule
	hideColNumber int // 隐藏的行

	weightColName      string // 重量
	weightColNameIndex int
	textureColName      string // 材质
	textureColNameIndex int
	workStageColName      string // 工序
	workStageColNameIndex int
	numColName      string // 数量
	numColNameIndex int
	workTimeColName      string // 工时
	workTimeColNameIndex int

	endColName      string // 最后的行
	endColNameIndex int

	fee1ColName      string // 钢材费用
	fee1ColNameIndex int

	fee2ColName      string // 工时费用
	fee2ColNameIndex int
	fee2ColNumber    int
}

const (
	TitleRowsNumber = 2 // 标题的行数
	DataSheetName   = "Sheet1"
	ModelSheetName  = "Sheet2"
)

func NewFee(f *excelize.File, isDebug bool) *Fee {
	return &Fee{
		f: f,
		isDebug: isDebug,
	}
}

func (l *Fee) Start() error {
	fmt.Println("2.检测文件是否完整")
	if err := l.check(); err != nil {
		fmt.Printf( "ERROR:文件读取错误,%v", err)
		return err
	}

	fmt.Println("3.计算耗材费用")
	if err := l.step1(); err != nil {
		fmt.Printf("ERROR:计算耗材费用错误,%v", err)
		return err
	}

	fmt.Println("4.计算工序费用")
	if err := l.step2(); err != nil {
		fmt.Printf("ERROR:计算工序费用错误,%v", err)
		return err
	}

	fmt.Println("5.汇总数据")
	if err := l.step3(); err != nil {
		fmt.Printf("ERROR:汇总数据错误,%v", err)
		return err
	}

	return nil
}

func (l *Fee) check() error {
	// 检测数据
	l.f.NewSheet(DataSheetName)
	rowsResp, err := l.f.GetRows(DataSheetName)
	if err != nil {
		return err
	}
	l.rows = rowsResp
	if len(rowsResp) < 2 {
		return errors.New("数据行数太少")
	}

	titleRows := rowsResp[TitleRowsNumber-1]
	for i, v := range titleRows {
		switch v {
		case "重量":
			l.weightColNameIndex = i + 1
			l.weightColName, _ = excelize.ColumnNumberToName(l.weightColNameIndex)
		case "材质":
			l.textureColNameIndex = i + 1
			l.textureColName, _ = excelize.ColumnNumberToName(l.textureColNameIndex)
		case "工序":
			l.workStageColNameIndex = i + 1
			l.workStageColName, _ = excelize.ColumnNumberToName(l.workStageColNameIndex)
		case "数量":
			l.numColNameIndex = i + 1
			l.numColName, _ = excelize.ColumnNumberToName(l.numColNameIndex)
		case "工时":
			l.workTimeColNameIndex = i + 1
			l.workTimeColName, _ = excelize.ColumnNumberToName(l.workTimeColNameIndex)
		}
	}
	l.endColNameIndex = len(titleRows)
	l.endColName, _ = excelize.ColumnNumberToName(l.endColNameIndex)

	if len(l.endColName) == 0 || len(l.weightColName) == 0 || len(l.textureColName) == 0 || len(l.workStageColName) == 0 || len(l.numColName) == 0 || len(l.workTimeColName) == 0 {
		return errors.New("检测数据缺少必要的行")
	}

	// 检测模板
	l.f.NewSheet(ModelSheetName)
	modeRowsResp, err := l.f.GetRows(ModelSheetName)
	if err != nil {
		return err
	}
	modeRows := make(map[string][]Rule, 0)
	root := ""
	for i, v := range modeRowsResp[0]  {
		if v != "" {
			root = v
			modeRows[root] = make([]Rule, 0)
		}
		tc := ""
		if len(modeRowsResp[2]) > i {
			tc = modeRowsResp[2][i]
		}
		tn := ""
		if len(modeRowsResp[1]) > i {
			tn = modeRowsResp[1][i]
		}
		modeRows[root] = append(modeRows[root], Rule{
			Name:    tn,
			PeopleContent: tc,
		})
		if len(modeRows[root]) > l.hideColNumber {
			l.hideColNumber = len(modeRows[root])
		}
	}
	l.modeRows = modeRows
	return nil
}

func (l *Fee) step1() error {
	l.fee1ColNameIndex = l.endColNameIndex + 1
	l.fee1ColName, _ = excelize.ColumnNumberToName(l.fee1ColNameIndex)

	// 标题
	l.f.NewSheet(DataSheetName)
	err := l.f.SetCellValue(DataSheetName, fmt.Sprintf("%v%v", l.fee1ColName, TitleRowsNumber), "钢材成本")
	if err != nil {
		return err
	}
	// 数据
	for i, v := range l.rows {
		if i <= TitleRowsNumber {
			continue
		}
		value := ""
		texture := v[l.textureColNameIndex-1]
		if strings.Contains(texture, "不锈钢") {
			value = fmt.Sprintf("=%v*%v%v*%v%v", 24.5, l.weightColName, strconv.FormatInt(int64(i+1), 10),l.numColName, strconv.FormatInt(int64(i+1), 10))
		} else if strings.Contains(texture, "铝") {
			value = fmt.Sprintf("=%v*%v%v*%v%v", 30, l.weightColName, strconv.FormatInt(int64(i+1), 10), l.numColName, strconv.FormatInt(int64(i+1), 10))
		} else {
			value = fmt.Sprintf("=%v*%v%v*%v%v", 7.5, l.weightColName,strconv.FormatInt(int64(i+1), 10), l.numColName, strconv.FormatInt(int64(i+1), 10))
		}
		err := l.f.SetCellFormula(DataSheetName, fmt.Sprintf("%v%v", l.fee1ColName, strconv.FormatInt(int64(i+1), 10)), value)
		if err != nil {
			return err
		}
	}
	return nil
}

func (l *Fee) rule(rowRule string, workStage, row int) string {
	rowRule = strings.ReplaceAll(rowRule, "重量", fmt.Sprintf("%v%v", l.weightColName, row))
	rowRule = strings.ReplaceAll(rowRule, "数量", fmt.Sprintf("%v%v", l.numColName, row))
	rowRule = strings.ReplaceAll(rowRule, "钢材成本", fmt.Sprintf("%v%v", l.fee1ColName, row))

	temp, _ := excelize.ColumnNumberToName(l.fee2ColNameIndex + workStage * (l.hideColNumber + 3) + 1)
	rowRule = strings.ReplaceAll(rowRule, "工时", fmt.Sprintf("%v%v", temp, row))
	return "=" + rowRule
}

func (l *Fee) step2() error {
	l.fee2ColNameIndex = l.fee1ColNameIndex + 1
	l.fee2ColName, _ = excelize.ColumnNumberToName(l.fee2ColNameIndex)
	l.fee2ColNumber = 0

	// 统计数据
	for j, v := range l.rows {
		if j < TitleRowsNumber {
			continue
		}
		workStageStr := v[l.workStageColNameIndex-1]
		if len(strings.Split(workStageStr, ",")) > l.fee2ColNumber {
			l.fee2ColNumber = len(strings.Split(workStageStr, ","))
		}
	}

	// 标题
	l.f.NewSheet(DataSheetName)
	for i := 0; i < l.fee2ColNumber; i++ {
		temp, _ := excelize.ColumnNumberToName(l.fee2ColNameIndex + i * (l.hideColNumber + 3))
		err := l.f.SetCellValue(DataSheetName, fmt.Sprintf("%v%v", temp, TitleRowsNumber), fmt.Sprintf("工序%v", i+1))
		if err != nil {
			return err
		}

		temp, _ = excelize.ColumnNumberToName(l.fee2ColNameIndex + i * (l.hideColNumber + 3) + 1)
		err = l.f.SetCellValue(DataSheetName, fmt.Sprintf("%v%v", temp, TitleRowsNumber), fmt.Sprintf("工时%v", i+1))
		if err != nil {
			return err
		}

		temp, _ = excelize.ColumnNumberToName(l.fee2ColNameIndex + i * (l.hideColNumber + 3) + 2)
		err = l.f.SetCellValue(DataSheetName, fmt.Sprintf("%v%v", temp, TitleRowsNumber), fmt.Sprintf("工序%v-费用", i+1))
		if err != nil {
			return err
		}

		for j := 0; j < l.hideColNumber; j ++ {
			temp, _ = excelize.ColumnNumberToName(l.fee2ColNameIndex + i * (l.hideColNumber + 3) + 3 + j)
			err = l.f.SetCellValue(DataSheetName, fmt.Sprintf("%v%v", temp, TitleRowsNumber), fmt.Sprintf("工序%v-费用-子类%v", i+1, j))
			if err != nil {
				return err
			}
		}

		if !l.isDebug {
			startPoint, _ := excelize.ColumnNumberToName(l.fee2ColNameIndex + i * (l.hideColNumber + 3) + 3)
			endPoint, _ := excelize.ColumnNumberToName(l.fee2ColNameIndex + i * (l.hideColNumber + 3) + 3 + l.hideColNumber - 1)
			err = l.f.SetColVisible(DataSheetName, fmt.Sprintf("%v:%v", startPoint, endPoint), false)
		}
	}

	// 数据
	for j, v := range l.rows {
		if j < TitleRowsNumber {
			continue
		}
		workStageStr := v[l.workStageColNameIndex-1]
		workTimeStr := v[l.workTimeColNameIndex-1]
		for i, value := range strings.Split(workStageStr, ",") {
			if value == "" {
				continue
			}
			temp, _ := excelize.ColumnNumberToName(l.fee2ColNameIndex + i * (l.hideColNumber + 3))
			err := l.f.SetCellValue(DataSheetName, fmt.Sprintf("%v%v", temp, j + 1), value)
			if err != nil {
				return err
			}

			// 计算
			value = strings.Trim(value, " ")
			if rvs, ok := l.modeRows[value]; ok {
				for rt, rv := range rvs {
					temp, _ = excelize.ColumnNumberToName(l.fee2ColNameIndex + i * (l.hideColNumber + 3) + 3 + rt)
					err = l.f.SetCellFormula(DataSheetName, fmt.Sprintf("%v%v", temp, j + 1), l.rule(rv.PeopleContent, i, j + 1))
					if err != nil {
						return err
					}
				}

				temp, _ = excelize.ColumnNumberToName(l.fee2ColNameIndex + i * (l.hideColNumber + 3) + 2)
				startPoint, _ := excelize.ColumnNumberToName(l.fee2ColNameIndex + i * (l.hideColNumber + 3) + 3)
				endPoint, _ := excelize.ColumnNumberToName(l.fee2ColNameIndex + i * (l.hideColNumber + 3) + 3 + l.hideColNumber - 1)
				err = l.f.SetCellFormula(DataSheetName, fmt.Sprintf("%v%v", temp, j + 1), fmt.Sprintf("=SUM(%v%v:%v%v)", startPoint, j + 1, endPoint, j + 1))
			} else {
				if !InArray(value, l.isAll) {
					l.isAll = append(l.isAll, value)
				}
			}
		}
		for i, value := range strings.Split(workTimeStr, ",") {
			temp, _ := excelize.ColumnNumberToName(l.fee2ColNameIndex + i * (l.hideColNumber + 3) + 1)
			err := l.f.SetCellValue(DataSheetName, fmt.Sprintf("%v%v", temp, j + 1), value)
			if err != nil {
				return err
			}
		}
	}
	if len(l.isAll) > 0 {
		fmt.Printf("存在不匹配的工序规则-%v\n", l.isAll)
	}

	return nil
}

func InArray(need interface{}, haystack interface{}) bool {
	switch key := need.(type) {
	case int:
		for _, item := range haystack.([]int) {
			if item == key {
				return true
			}
		}
	case string:
		for _, item := range haystack.([]string) {
			if item == key {
				return true
			}
		}
	case int64:
		for _, item := range haystack.([]int64) {
			if item == key {
				return true
			}
		}
	case float64:
		for _, item := range haystack.([]float64) {
			if item == key {
				return true
			}
		}
	default:
		return false
	}
	return false
}


func (l *Fee) step3() error {
	// 标题
	endTemp, _ := excelize.ColumnNumberToName(l.fee2ColNameIndex + (l.hideColNumber + 3) * l.fee2ColNumber + 1)
	err := l.f.SetCellValue(DataSheetName, fmt.Sprintf("%v%v", endTemp, TitleRowsNumber), fmt.Sprintf("合计"))
	if err != nil {
		return err
	}

	for i, _ := range l.rows {
		if i <= TitleRowsNumber {
			continue
		}

		value := make([]string, 0)
		for j := 0; j < l.fee2ColNumber; j++ {
			cTemp, _ := excelize.ColumnNumberToName(l.fee2ColNameIndex + (l.hideColNumber + 3) * j + 2)
			value = append(value, fmt.Sprintf("%v%v", cTemp, i))
		}
		value = append(value, fmt.Sprintf("%v%v", l.fee1ColName, i))

		err := l.f.SetCellFormula(DataSheetName, fmt.Sprintf("%v%v", endTemp, i), fmt.Sprintf("=SUM(%v)", strings.Join(value, ",")))
		if err != nil {
			return err
		}
	}

	// 字体设置黑色
	style, err := l.f.NewStyle(`{"font":{"bold":true}}`)
	_ = l.f.SetCellStyle(DataSheetName, fmt.Sprintf("%v%v", "A", TitleRowsNumber), fmt.Sprintf("%v%v", endTemp, TitleRowsNumber), style)

	// 设置背景颜色
	style, _ = l.f.NewStyle(`{"fill":{"type":"pattern","color":["#d67e4b"],"pattern":1}}`)
	_ = l.f.SetCellStyle(DataSheetName, fmt.Sprintf("%v%v", endTemp, 1), fmt.Sprintf("%v%v", endTemp, len(l.rows)), style)
	_ = l.f.SetCellStyle(DataSheetName, fmt.Sprintf("%v%v", l.fee1ColName, 1), fmt.Sprintf("%v%v", l.fee1ColName, len(l.rows)), style)

	for j := 0; j < l.fee2ColNumber; j++ {
		cTemp, _ := excelize.ColumnNumberToName(l.fee2ColNameIndex + (l.hideColNumber + 3) * j + 2)
		_ = l.f.SetCellStyle(DataSheetName, fmt.Sprintf("%v%v", cTemp, 1), fmt.Sprintf("%v%v", cTemp, len(l.rows)), style)
	}
	return nil
}
