package main

import (
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

const (
	outputName = "对比结果.xlsx"
	colCount   = 6 // A~F
	firstCell  = 3 // 起始数据行（第 1 行放文件名、第 2 行放表头）
)

// bomRow 表示 BOM 中的一行，固定保存 A~F 六列的值。
type bomRow struct {
	cells [colCount]string
}

func (r bomRow) isEmpty() bool {
	for _, v := range r.cells {
		if strings.TrimSpace(v) != "" {
			return false
		}
	}
	return true
}

// normQty 归一化数量：能解析成数字时按数值比较（1 与 1.0、01 视为相同），否则按去空格后的小写文本比较。
func normQty(s string) string {
	s = strings.TrimSpace(s)
	if s == "" {
		return ""
	}
	clean := strings.ReplaceAll(s, ",", "")
	if f, err := strconv.ParseFloat(clean, 64); err == nil {
		return strconv.FormatFloat(f, 'f', -1, 64)
	}
	return strings.ToLower(s)
}

// matchKey 料号(A列) + 数量(E列) 共同构成匹配键。
func (r bomRow) matchKey() string {
	return strings.TrimSpace(r.cells[0]) + "\x1f" + normQty(r.cells[4])
}

// readBOM 读取 BOM 的第一个工作表：返回表头行与 A~F 的有效数据行（跳过整行空白）。
func readBOM(path string) (bomRow, []bomRow, error) {
	var header bomRow
	f, err := excelize.OpenFile(path)
	if err != nil {
		return header, nil, fmt.Errorf("无法打开文件 %s：%v", filepath.Base(path), err)
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return header, nil, fmt.Errorf("文件 %s 中没有工作表", filepath.Base(path))
	}

	rows, err := f.GetRows(sheets[0])
	if err != nil {
		return header, nil, fmt.Errorf("读取 %s 出错：%v", filepath.Base(path), err)
	}

	if len(rows) == 0 {
		return header, nil, fmt.Errorf("文件 %s 是空表，缺少表头", filepath.Base(path))
	}
	for i := 0; i < colCount && i < len(rows[0]); i++ {
		header.cells[i] = strings.TrimSpace(rows[0][i])
	}

	var data []bomRow
	for i := 1; i < len(rows); i++ { // 跳过表头
		var r bomRow
		for c := 0; c < colCount && c < len(rows[i]); c++ {
			r.cells[c] = strings.TrimSpace(rows[i][c])
		}
		if r.isEmpty() {
			continue // 整行空白不参与比较
		}
		data = append(data, r)
	}
	return header, data, nil
}

// pickInputs 决定要比较的两个文件：优先使用命令行参数，否则自动扫描当前文件夹里的 xlsx。
func pickInputs(args []string) (string, string, error) {
	if len(args) == 2 {
		return args[0], args[1], nil
	}
	if len(args) > 0 && len(args) != 2 {
		return "", "", fmt.Errorf("需要 2 个文件，实际给了 %d 个。用法：bom-compare.exe [BOM1.xlsx BOM2.xlsx]", len(args))
	}

	files, err := filepath.Glob("*.xlsx")
	if err != nil {
		return "", "", fmt.Errorf("扫描当前文件夹失败：%v", err)
	}
	var cands []string
	for _, p := range files {
		if strings.HasPrefix(filepath.Base(p), "~$") { // 跳过 Excel 打开时产生的临时文件
			continue
		}
		if strings.EqualFold(filepath.Base(p), outputName) {
			continue // 跳过上一次的结果文件
		}
		cands = append(cands, p)
	}
	sort.Strings(cands)

	if len(cands) != 2 {
		if len(cands) == 0 {
			return "", "", fmt.Errorf("当前文件夹里没有找到 xlsx 文件")
		}
		return "", "", fmt.Errorf("当前文件夹里有 %d 个 xlsx 文件（%s），请用命令行参数指定要比较的两个，例如：bom-compare.exe BOM1.xlsx BOM2.xlsx",
			len(cands), strings.Join(baseNames(cands), "、"))
	}
	return cands[0], cands[1], nil
}

func baseNames(paths []string) []string {
	out := make([]string, len(paths))
	for i, p := range paths {
		out[i] = filepath.Base(p)
	}
	return out
}

func main() {
	exitCode := run()
	os.Exit(exitCode)
}

func run() int {
	defer enableUTF8Console()()

	p1, p2, err := pickInputs(os.Args[1:])
	if err != nil {
		fmt.Println("错误：" + err.Error())
		return 2
	}

	header1, data1, err := readBOM(p1)
	if err != nil {
		fmt.Println("错误：" + err.Error())
		return 1
	}
	header2, data2, err := readBOM(p2)
	if err != nil {
		fmt.Println("错误：" + err.Error())
		return 1
	}

	fmt.Printf("第1个BOM：%s（数据 %d 行）\n", filepath.Base(p1), len(data1))
	fmt.Printf("第2个BOM：%s（数据 %d 行）\n", filepath.Base(p2), len(data2))

	// 按「料号+数量」为第2个BOM建索引；同键重复时按出现顺序先进先出配对。
	idx2 := make(map[string][]int, len(data2))
	for i := range data2 {
		k := data2[i].matchKey()
		idx2[k] = append(idx2[k], i)
	}

	used2 := make([]bool, len(data2))
	type pair struct{ i1, i2 int }
	var matched []pair
	var only1, only2 []int

	for i := range data1 {
		k := data1[i].matchKey()
		q := idx2[k]
		if len(q) > 0 {
			j := q[0]
			idx2[k] = q[1:]
			used2[j] = true
			matched = append(matched, pair{i, j})
		} else {
			only1 = append(only1, i)
		}
	}
	for j := range data2 {
		if !used2[j] {
			only2 = append(only2, j)
		}
	}

	// 每次运行都重建结果文件：先删旧的。
	if err := os.Remove(outputName); err != nil && !os.IsNotExist(err) {
		fmt.Printf("错误：无法删除已存在的 %s（可能正被 Excel 打开）：%v\n", outputName, err)
		return 1
	}

	f := excelize.NewFile()
	defer f.Close()
	sheet := "对比结果"
	f.SetSheetName("Sheet1", sheet)

	writeRow := func(rowNum int, startCol int, r bomRow) {
		for c := 0; c < colCount; c++ {
			if v := r.cells[c]; v != "" {
				cell, _ := excelize.CoordinatesToCellName(startCol+c, rowNum)
				f.SetCellStr(sheet, cell, v)
			}
		}
	}

	// 第1行：两个 BOM 的文件名；第2行：表头。
	f.SetCellStr(sheet, "A1", filepath.Base(p1))
	f.SetCellStr(sheet, "G1", filepath.Base(p2))
	writeRow(2, 1, header1)
	h2 := header2
	for c := 0; c < colCount; c++ {
		if h2.cells[c] == "" {
			h2.cells[c] = header1.cells[c] // 第2个BOM表头缺失时沿用第1个
		}
	}
	writeRow(2, 7, h2)

	rowNum := firstCell
	for _, p := range matched {
		writeRow(rowNum, 1, data1[p.i1])
		writeRow(rowNum, 7, data2[p.i2])
		f.SetCellStr(sheet, fmt.Sprintf("M%d", rowNum), "共有物料")
		rowNum++
	}
	for _, i := range only1 {
		writeRow(rowNum, 1, data1[i])
		f.SetCellStr(sheet, fmt.Sprintf("M%d", rowNum), "第1个BOM特有物料")
		rowNum++
	}
	for _, j := range only2 {
		writeRow(rowNum, 7, data2[j])
		f.SetCellStr(sheet, fmt.Sprintf("M%d", rowNum), "第2个BOM特有物料")
		rowNum++
	}

	styleSheet(f, sheet, rowNum-1)

	// 保存：覆盖为新文件。
	if err := f.SaveAs(outputName); err != nil {
		fmt.Printf("错误：保存 %s 失败：%v\n", outputName, err)
		return 1
	}

	fmt.Println()
	fmt.Printf("完成：共有物料 %d 行，第1个BOM特有 %d 行，第2个BOM特有 %d 行。\n", len(matched), len(only1), len(only2))
	fmt.Printf("结果已写入：%s（当前文件夹）\n", outputName)
	return 0
}

// styleSheet 做基础排版：表头加粗居中、灰底、设列宽、冻结前两行。
func styleSheet(f *excelize.File, sheet string, lastRow int) {
	headerStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"F2F2F2"}},
	})
	labelStyle, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	f.SetCellStyle(sheet, "A1", "M2", headerStyle)
	if lastRow >= 3 {
		f.SetCellStyle(sheet, "M3", fmt.Sprintf("M%d", lastRow), labelStyle)
	}

	f.SetColWidth(sheet, "A", "A", 18)
	f.SetColWidth(sheet, "B", "F", 14)
	f.SetColWidth(sheet, "G", "G", 18)
	f.SetColWidth(sheet, "H", "L", 14)
	f.SetColWidth(sheet, "M", "M", 20)

	f.SetPanes(sheet, &excelize.Panes{
		Freeze:      true,
		YSplit:      2,
		TopLeftCell: "A3",
		ActivePane:  "bottomLeft",
	})
}
//（注：内容由AI生成）
