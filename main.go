package main

import (
	"fmt"
	"os"
	"path/filepath"
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
	ColE string // 数量
	ColF string
	Key  string // A料号+E数量 作为匹配key
}

func makeKey(a, e string) string {
	return fmt.Sprintf("%s|%s", strings.TrimSpace(a), strings.TrimSpace(e))
}

func readBOM(path string) ([]BomRow, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	sheet := f.GetSheetList()[0]
	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}

	var list []BomRow
	// 跳过第一行表头，从第2行开始读数据
	for i := 1; i < len(rows); i++ {
		r := rows[i]
		a := getCell(r, 0)
		b := getCell(r, 1)
		c := getCell(r, 2)
		d := getCell(r, 3)
		e := getCell(r, 4)
		ff := getCell(r, 5)
		key := makeKey(a, e)
		list = append(list, BomRow{
			ColA: a, ColB: b, ColC: c, ColD: d, ColE: e, ColF: ff, Key: key,
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

func main() {
	enableUTF8Console()
	args := os.Args[1:]
	var file1, file2 string
	dir, _ := os.Getwd()

	if len(args) >= 2 {
		file1 = args[0]
		file2 = args[1]
	} else {
		// 自动查找当前目录xlsx，排除对比结果.xlsx 和临时文件
		var xlsxFiles []string
		_ = filepath.Walk(dir, func(p string, info os.FileInfo, err error) error {
			if err != nil {
				return nil
			}
			if info.IsDir() {
				return nil
			}
			name := info.Name()
			if strings.HasSuffix(strings.ToLower(name), ".xlsx") {
				if name == "对比结果.xlsx" || strings.HasPrefix(name, "~$") {
					return nil
				}
				xlsxFiles = append(xlsxFiles, p)
			}
			return nil
		})
		if len(xlsxFiles) < 2 {
			fmt.Println("错误：当前目录需要至少2个xlsx BOM文件，或者手动传入参数 bom-compare.exe 1.xlsx 2.xlsx")
			os.Exit(1)
		}
		file1 = xlsxFiles[0]
		file2 = xlsxFiles[1]
	}

	// 读取两个BOM
	list1, err := readBOM(file1)
	if err != nil {
		fmt.Printf("读取%s失败：%v\n", file1, err)
		os.Exit(1)
	}
	list2, err := readBOM(file2)
	if err != nil {
		fmt.Printf("读取%s失败：%v\n", file2, err)
		os.Exit(1)
	}

	name1 := filepath.Base(file1)
	name2 := filepath.Base(file2)

	// 建立list2 key -> 剩余行索引
	type idxWrap struct {
		used bool
		row  BomRow
	}
	map2 := make(map[string][]*idxWrap)
	for _, r := range list2 {
		map2[r.Key] = append(map2[r.Key], &idxWrap{used: false, row: r})
	}

	var common []struct{ r1, r2 BomRow }
	var only1 []BomRow
	var only2 []BomRow

	// 匹配共有物料
	for _, r1 := range list1 {
		found := false
		if arr, ok := map2[r1.Key]; ok {
			for _, item := range arr {
				if !item.used {
					common = append(common, struct{ r1, r2 BomRow }{r1, item.row})
					item.used = true
					found = true
					break
				}
			}
		}
		if !found {
			only1 = append(only1, r1)
		}
	}
	// 收集list2剩下没匹配上的，属于only2
	for _, arr := range map2 {
		for _, item := range arr {
			if !item.used {
				only2 = append(only2, item.row)
			}
		}
	}

	// 删除旧结果文件
	outName := "对比结果.xlsx"
	_ = os.Remove(outName)

	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetSheetName("Sheet1", sheet)

	// A1=文件1名，G1=文件2名
	f.SetCellValue(sheet, "A1", name1)
	f.SetCellValue(sheet, "G1", name2)

	// 第2行表头 A-F / G-L
	headers := []string{"料号", "B", "C", "D", "数量", "F", "料号", "B", "C", "D", "数量", "F", "类型"}
	for colIdx, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(colIdx+1, 2)
		f.SetCellValue(sheet, cell, h)
	}

	row := 3 // 数据从第3行开始
	// 写入共有物料
	for _, item := range common {
		writeRow(f, sheet, row, item.r1, item.r2, "共有物料")
		row++
	}
	// 写入BOM1特有
	for _, r := range only1 {
		writeRow(f, sheet, row, r, BomRow{}, "第1个BOM特有物料")
		row++
	}
	// 写入BOM2特有
	for _, r := range only2 {
		writeRow(f, sheet, row, BomRow{}, r, "第2个BOM特有物料")
		row++
	}

	if err := f.SaveAs(outName); err != nil {
		fmt.Printf("保存对比结果.xlsx失败，请关闭打开的Excel！err:%v\n", err)
		os.Exit(1)
	}
	fmt.Printf("对比完成！输出文件：%s\n", outName)
}

func writeRow(f *excelize.File, sheet string, row int, r1, r2 BomRow, typ string) {
	// A-F r1
	f.SetCellValue(sheet, fmt.Sprintf("A%d", row), r1.ColA)
	f.SetCellValue(sheet, fmt.Sprintf("B%d", row), r1.ColB)
	f.SetCellValue(sheet, fmt.Sprintf("C%d", row), r1.ColC)
	f.SetCellValue(sheet, fmt.Sprintf("D%d", row), r1.ColD)
	f.SetCellValue(sheet, fmt.Sprintf("E%d", row), r1.ColE)
	f.SetCellValue(sheet, fmt.Sprintf("F%d", row), r1.ColF)

	// G-L r2
	f.SetCellValue(sheet, fmt.Sprintf("G%d", row), r2.ColA)
	f.SetCellValue(sheet, fmt.Sprintf("H%d", row), r2.ColB)
	f.SetCellValue(sheet, fmt.Sprintf("I%d", row), r2.ColC)
	f.SetCellValue(sheet, fmt.Sprintf("J%d", row), r2.ColD)
	f.SetCellValue(sheet, fmt.Sprintf("K%d", row), r2.ColE)
	f.SetCellValue(sheet, fmt.Sprintf("L%d", row), r2.ColF)

	// M列类型
	f.SetCellValue(sheet, fmt.Sprintf("M%d", row), typ)
}
