package main

import (
	"bufio"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"syscall"

	"github.com/xuri/excelize/v2"
)

var (
	kernel32               = syscall.NewLazyDLL("kernel32.dll")
	procSetConsoleOutputCP = kernel32.NewProc("SetConsoleOutputCP")
)

func enableUTF8Console() {
	const CP_UTF8 = 65001
	procSetConsoleOutputCP.Call(uintptr(CP_UTF8))
}

type BomRow struct {
	ColA string
	ColB string
	ColC string
	ColD string
	ColE string //数量
	ColF string //位号
}

// PriceMap 料号 -> 含税单价
var priceMap map[string]float64

func parseQty(s string) float64 {
	s = strings.TrimSpace(s)
	f, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return 0
	}
	return f
}

// 读取价格总表.xlsx，A列料号，C列含税单价
func loadPriceTable() error {
	priceMap = make(map[string]float64)
	priceFile := "价格总表.xlsx"
	if _, err := os.Stat(priceFile); err != nil {
		// 文件不存在，不报错，只是所有料号都查不到价格
		return nil
	}
	f, err := excelize.OpenFile(priceFile)
	if err != nil {
		return err
	}
	defer f.Close()
	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil
	}
	rows, err := f.GetRows(sheets[0])
	if err != nil {
		return err
	}
	for i := 1; i < len(rows); i++ {
		r := rows[i]
		pn := getCell(r, 0)
		priceStr := getCell(r, 2)
		pn = strings.TrimSpace(pn)
		if pn == "" {
			continue
		}
		pr := parseQty(priceStr)
		priceMap[pn] = pr
	}
	return nil
}

func getPrice(pn string) (float64, bool) {
	pn = strings.TrimSpace(pn)
	v, ok := priceMap[pn]
	return v, ok
}

// readBOM 增强：跳过合并单元格非首行（E、F列合并区域）
func readBOM(path string) ([]BomRow, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("文件无工作表")
	}
	sheet := sheets[0]

	// 获取所有合并单元格
	mergeRanges, err := f.GetMergeCells(sheet)
	if err != nil {
		return nil, err
	}
	// key: 行号（从0开始）, value:true 代表这一行属于某个合并区域，但不是首行，需要跳过
	skipRow := make(map[int]bool)
	for _, mr := range mergeRanges {
		startCol, startRow, endCol, endRow, err := excelize.CellNameToCoordinates(mr)
		if err != nil {
			continue
		}
		// Excel坐标是从1开始，转为程序内0基
		sR := startRow - 1
		eR := endRow - 1
		sC := startCol
		eC := endCol
		// 判断这个合并区域是否包含E(5)或者F(6)列
		hasE := (sC <= 5 && eC >= 5)
		hasF := (sC <= 6 && eC >= 6)
		if hasE || hasF {
			// 合并区内，除首行外全部标记跳过
			for r := sR + 1; r <= eR; r++ {
				skipRow[r] = true
			}
		}
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}
	var list []BomRow
	for i := 1; i < len(rows); i++ {
		if skipRow[i] {
			continue
		}
		r := rows[i]
		a := getCell(r, 0)
		b := getCell(r, 1)
		c := getCell(r, 2)
		d := getCell(r, 3)
		e := getCell(r, 4)
		ff := getCell(r, 5)
		list = append(list, BomRow{
			ColA: a, ColB: b, ColC: c, ColD: d, ColE: e, ColF: ff,
		})
	}
	return list, nil
}

func getCell(row []string, idx int) string {
	if idx >= len(row) {
		return ""
	}
	return strings.TrimSpace(row[idx])
}

type WriteRowData struct {
	r1  BomRow
	r2  BomRow
	typ string
}

// sortGroup：一组WriteRowData，内部排序：查到价格在前，无价格在后
func sortGroup(list []WriteRowData) []WriteRowData {
	out := make([]WriteRowData, 0, len(list))
	hasPrice := make([]WriteRowData, 0)
	noPrice := make([]WriteRowData, 0)
	for _, d := range list {
		// 判断规则：只要r1料号查到 或者 r2料号查到，不算无价格？
		// 需求：该行【对应的】料号查不到就标记无价格
		// 共有物料：看r1料号；BOM1特有：r1料号；BOM2特有：r2料号
		no := false
		switch d.typ {
		case "共有物料", "第1个BOM特有物料":
			_, ok := getPrice(d.r1.ColA)
			if !ok {
				no = true
			}
		case "第2个BOM特有物料":
			_, ok := getPrice(d.r2.ColA)
			if !ok {
				no = true
			}
		}
		if no {
			noPrice = append(noPrice, d)
		} else {
			hasPrice = append(hasPrice, d)
		}
	}
	out = append(out, hasPrice...)
	out = append(out, noPrice...)
	return out
}

func main() {
	enableUTF8Console()
	// 加载价格表
	if err := loadPriceTable(); err != nil {
		fmt.Printf("读取价格总表.xlsx异常：%v\n", err)
		waitEnter()
		os.Exit(1)
	}

	args := os.Args[1:]
	var file1, file2 string
	dir, _ := os.Getwd()

	if len(args) >= 2 {
		file1 = args[0]
		file2 = args[1]
	} else {
		var xlsxFiles []string
		_ = filepath.Walk(dir, func(p string, info os.FileInfo, err error) error {
			if err != nil {
				return nil
			}
			if info.IsDir() {
				return nil
			}
			name := info.Name()
			lower := strings.ToLower(name)
			if strings.HasSuffix(lower, ".xlsx") {
				if name == "对比结果.xlsx" || name == "价格总表.xlsx" || strings.HasPrefix(name, "~$") {
					return nil
				}
				xlsxFiles = append(xlsxFiles, p)
			}
			return nil
		})
		if len(xlsxFiles) < 2 {
			fmt.Println("错误：当前目录需要至少2个xlsx BOM文件，或者手动传入参数 bom-compare.exe 1.xlsx 2.xlsx")
			waitEnter()
			os.Exit(1)
		}
		file1 = xlsxFiles[0]
		file2 = xlsxFiles[1]
	}

	list1, err := readBOM(file1)
	if err != nil {
		fmt.Printf("读取%s失败：%v\n", file1, err)
		waitEnter()
		os.Exit(1)
	}
	list2, err := readBOM(file2)
	if err != nil {
		fmt.Printf("读取%s失败：%v\n", file2, err)
		waitEnter()
		os.Exit(1)
	}

	name1 := filepath.Base(file1)
	name2 := filepath.Base(file2)

	type wrap struct {
		used bool
		row  BomRow
	}
	map2 := make(map[string][]*wrap)
	for _, r := range list2 {
		pn := strings.TrimSpace(r.ColA)
		map2[pn] = append(map2[pn], &wrap{used: false, row: r})
	}

	var only1NoMatch []BomRow
	var only2NoMatch []BomRow
	type pair struct {
		r1 BomRow
		r2 BomRow
	}
	var commonPairs []pair

	for _, r1 := range list1 {
		pn1 := strings.TrimSpace(r1.ColA)
		arr, ok := map2[pn1]
		if !ok {
			only1NoMatch = append(only1NoMatch, r1)
			continue
		}
		foundIdx := -1
		for i, item := range arr {
			if !item.used {
				foundIdx = i
				break
			}
		}
		if foundIdx == -1 {
			only1NoMatch = append(only1NoMatch, r1)
			continue
		}
		item := arr[foundIdx]
		item.used = true
		commonPairs = append(commonPairs, pair{r1, item.row})
	}

	for _, arr := range map2 {
		for _, w := range arr {
			if !w.used {
				only2NoMatch = append(only2NoMatch, w.row)
			}
		}
	}

	var commonList []WriteRowData
	var bom1ExtraList []WriteRowData
	var bom2ExtraList []WriteRowData

	for _, p := range commonPairs {
		r1 := p.r1
		r2 := p.r2
		q1 := parseQty(r1.ColE)
		q2 := parseQty(r2.ColE)
		minQty := math.Min(q1, q2)
		diff := math.Abs(q1 - q2)

		sharedR1 := r1
		sharedR1.ColE = fmt.Sprintf("%g", minQty)
		sharedR2 := r2
		sharedR2.ColE = fmt.Sprintf("%g", minQty)
		commonList = append(commonList, WriteRowData{sharedR1, sharedR2, "共有物料"})

		if diff > 1e-9 {
			if q1 > q2 {
				diffRow := r1
				diffRow.ColE = fmt.Sprintf("%g", diff)
				bom1ExtraList = append(bom1ExtraList, WriteRowData{diffRow, BomRow{}, "第1个BOM特有物料"})
			} else {
				diffRow := r2
				diffRow.ColE = fmt.Sprintf("%g", diff)
				bom2ExtraList = append(bom2ExtraList, WriteRowData{BomRow{}, diffRow, "第2个BOM特有物料"})
			}
		}
	}

	for _, r := range only1NoMatch {
		bom1ExtraList = append(bom1ExtraList, WriteRowData{r, BomRow{}, "第1个BOM特有物料"})
	}
	for _, r := range only2NoMatch {
		bom2ExtraList = append(bom2ExtraList, WriteRowData{BomRow{}, r, "第2个BOM特有物料"})
	}

	// 每组内部排序：有价格在前，无价格在后
	commonList = sortGroup(commonList)
	bom1ExtraList = sortGroup(bom1ExtraList)
	bom2ExtraList = sortGroup(bom2ExtraList)

	outName := "对比结果.xlsx"
	_ = os.Remove(outName)
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetSheetName("Sheet1", sheet)

	f.SetCellValue(sheet, "A1", name1)
	f.SetCellValue(sheet, "I1", name2)

	// 新表头！
	headers := []string{
		"料号", "B", "C", "D", "数量", "F",
		"单价", "总价",
		"料号", "B", "C", "D", "数量", "F",
		"单价", "总价",
		"类型",
	}
	for colIdx, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(colIdx+1, 2)
		f.SetCellValue(sheet, cell, h)
	}

	row := 3
	for _, d := range commonList {
		writeRowNew(f, sheet, row, d.r1, d.r2, d.typ)
		row++
	}
	for _, d := range bom1ExtraList {
		writeRowNew(f, sheet, row, d.r1, d.r2, d.typ)
		row++
	}
	for _, d := range bom2ExtraList {
		writeRowNew(f, sheet, row, d.r1, d.r2, d.typ)
		row++
	}

	if err := f.SaveAs(outName); err != nil {
		fmt.Printf("保存对比结果.xlsx失败，请关闭打开的Excel！err:%v\n", err)
		waitEnter()
		os.Exit(1)
	}

	fmt.Println("已经对比完成，对比结果请查收当前文件夹中的对比结果.xlsx")
	waitEnter()
}

// writeRowNew 新版写入，支持单价总价
func writeRowNew(f *excelize.File, sheet string, row int, r1, r2 BomRow, typ string) {
	// A-F：BOM1基础
	f.SetCellValue(sheet, fmt.Sprintf("A%d", row), r1.ColA)
	f.SetCellValue(sheet, fmt.Sprintf("B%d", row), r1.ColB)
	f.SetCellValue(sheet, fmt.Sprintf("C%d", row), r1.ColC)
	f.SetCellValue(sheet, fmt.Sprintf("D%d", row), r1.ColD)
	f.SetCellValue(sheet, fmt.Sprintf("E%d", row), r1.ColE)
	f.SetCellValue(sheet, fmt.Sprintf("F%d", row), r1.ColF)

	// G H：BOM1单价、总价
	p1, ok1 := getPrice(r1.ColA)
	q1 := parseQty(r1.ColE)
	if ok1 {
		f.SetCellValue(sheet, fmt.Sprintf("G%d", row), p1)
		f.SetCellValue(sheet, fmt.Sprintf("H%d", row), p1*q1)
	} else {
		f.SetCellValue(sheet, fmt.Sprintf("G%d", row), "价格不存在")
		f.SetCellValue(sheet, fmt.Sprintf("H%d", row), "价格不存在")
	}

	// I-N：BOM2基础
	f.SetCellValue(sheet, fmt.Sprintf("I%d", row), r2.ColA)
	f.SetCellValue(sheet, fmt.Sprintf("J%d", row), r2.ColB)
	f.SetCellValue(sheet, fmt.Sprintf("K%d", row), r2.ColC)
	f.SetCellValue(sheet, fmt.Sprintf("L%d", row), r2.ColD)
	f.SetCellValue(sheet, fmt.Sprintf("M%d", row), r2.ColE)
	f.SetCellValue(sheet, fmt.Sprintf("N%d", row), r2.ColF)

	// O P：BOM2单价、总价
	p2, ok2 := getPrice(r2.ColA)
	q2 := parseQty(r2.ColE)
	if ok2 {
		f.SetCellValue(sheet, fmt.Sprintf("O%d", row), p2)
		f.SetCellValue(sheet, fmt.Sprintf("P%d", row), p2*q2)
	} else {
		f.SetCellValue(sheet, fmt.Sprintf("O%d", row), "价格不存在")
		f.SetCellValue(sheet, fmt.Sprintf("P%d", row), "价格不存在")
	}

	// Q 类型
	f.SetCellValue(sheet, fmt.Sprintf("Q%d", row), typ)
}

func waitEnter() {
	fmt.Println("按回车键退出...")
	scanner := bufio.NewScanner(os.Stdin)
	scanner.Scan()
}
