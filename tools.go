package main

import (
	"bufio"
	"crypto/md5"
	"encoding/hex"
	"fmt"
	"io"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/liserjrqlxue/goUtil/osUtil"
	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/xuri/excelize/v2"
)

func ExcelToSlice(filename, sheetName string) ([][]string, error) {
	// Open the Excel file
	file, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, err
	}
	defer simpleUtil.DeferClose(file)

	// Get all the rows from the specified sheet
	return file.GetRows(sheetName)
}

func slice2MapArray(s [][]string) (data []map[string]string) {
	var key = s[0]
	for i := 1; i < len(s); i++ {
		var item = make(map[string]string)
		for j := 0; j < len(s[i]); j++ {
			item[key[j]] = s[i][j]
		}
		data = append(data, item)
	}
	return
}

func RunPrimerDesigner(workdir, name, seq string, args ...string) error {
	workdir = filepath.Join(workdir, name)
	seqFile := filepath.Join(workdir, "seq.txt")
	prefix := filepath.Join(workdir, name)

	if err := os.MkdirAll(workdir, 0755); err != nil {
		return err
	}

	if err := os.WriteFile(seqFile, []byte(seq), 0644); err != nil {
		return err
	}

	args = append(args, "-i", seqFile, "-o", prefix, "-n", name)

	cmd := exec.Command("util/primerDesigner/primerDesigner", args...)
	cmd.Stderr = os.Stderr
	cmd.Stdout = os.Stdout
	log.Println(cmd)

	return cmd.Run()
}

func writeBatch(prefix string, list [][3]string) error {
	var count = 0
	var panel = 0
	var fList []*os.File
	var TailChangingPrimers = osUtil.Create(prefix + "换尾引物.txt")
	defer simpleUtil.DeferClose(TailChangingPrimers)

	_, err := fmt.Fprintf(TailChangingPrimers, "%s\t%s\t%s\n", "引物名称", "引物序列5-3", "基因名")
	if err != nil {
		return err
	}

	excel, err := excelize.OpenFile(prefix + "J-" + list[0][0] + ".xlsx")
	if err != nil {
		return err
	}

	var (
		selfSyntheticExcel = excelize.NewFile()
		panelSheetName     = "Sheet1"
		panelName          = ""
		panelCount         = 0
		panelCol           = 3
		panelRow           = 0

		bgColor1   = "#DDEBF7"
		bgColor2   = "#FCE4D6"
		bgColor3   = "#E2EFDA"
		tableStyle = excelize.Style{
			Border: []excelize.Border{
				{Type: "left", Color: "000000", Style: 1},
				{Type: "bottom", Color: "000000", Style: 1},
				{Type: "right", Color: "000000", Style: 1},
				{Type: "bottom", Color: "000000", Style: 1},
			},
			Alignment: &excelize.Alignment{
				WrapText:   true,
				Horizontal: "center",
				Vertical:   "center",
			},
		}
		bgStyle1 = excelize.Style{
			Border:    tableStyle.Border,
			Alignment: tableStyle.Alignment,
			Fill: excelize.Fill{
				Type:    "pattern",
				Color:   []string{bgColor1},
				Pattern: 1,
			},
		}
		bgStyle2 = excelize.Style{
			Border:    tableStyle.Border,
			Alignment: tableStyle.Alignment,
			Fill: excelize.Fill{
				Type:    "pattern",
				Color:   []string{bgColor2},
				Pattern: 1,
			},
		}
		bgStyle3 = excelize.Style{
			Border:    tableStyle.Border,
			Alignment: tableStyle.Alignment,
			Fill: excelize.Fill{
				Type:    "pattern",
				Color:   []string{bgColor3},
				Pattern: 1,
			},
		}
	)
	tableStyleId, err := selfSyntheticExcel.NewStyle(&tableStyle)
	if err != nil {
		return err
	}
	bgStyle1Id, err := selfSyntheticExcel.NewStyle(&bgStyle1)
	if err != nil {
		return err
	}
	bgStyle2Id, err := selfSyntheticExcel.NewStyle(&bgStyle2)
	if err != nil {
		return err
	}
	bgStyle3Id, err := selfSyntheticExcel.NewStyle(&bgStyle3)
	if err != nil {
		return err
	}
	var bgStyleMap = map[int]int{
		0: bgStyle1Id,
		1: bgStyle2Id,
		2: bgStyle3Id,
	}

	for i, seq := range list {
		var bgStyleCurrent = bgStyleMap[i%3]
		id := seq[0]
		lines, err := ReadFileToLineArray(seq[1])
		if err != nil {
			return err
		}

		for _, line := range lines {
			if count%96 == 0 {
				var tag = string(rune('A' + panel))
				if panel > 25 {
					tag = string(rune('A'+panel-26)) + "_1"
				}
				panelName = filepath.Base(prefix) + tag + "-" + id
				log.Printf("%d:\t%s\n", panel, panelName)

				panelCount++
				panelRow = 1 + (panelCount-1)*11
				// fmt.Printf("%c%d:版号:%s\n", 'D', panelRow, panelName)
				err := CreatePanelTable(selfSyntheticExcel, panelSheetName, panelName, panelCol, panelRow, tableStyleId)
				if err != nil {
					log.Printf("CreatePanelTable [%s:%s %d,%d] fail:[%v]", panelName, panelSheetName, panelCol, panelRow, err)
					return err
				}
				panelNameCell, err := excelize.CoordinatesToCellName(panelCol+1, panelRow)
				if err != nil {
					return err
				}
				err = selfSyntheticExcel.SetCellRichText(
					panelSheetName,
					panelNameCell,
					[]excelize.RichTextRun{
						{
							Text: filepath.Base(prefix) + tag,
							Font: &excelize.Font{
								Color: "#FF0000",
							},
						},
						{
							Text: "-" + id,
						},
					},
				)
				if err != nil {
					return err
				}

				f, err := os.Create(prefix + tag + "-" + id + ".seq")
				if err != nil {
					return err
				}
				panel++
				fList = append(fList, f)
			}
			_, err = fList[panel-1].WriteString(line + "\n")
			if err != nil {
				return err
			}
			var cells = strings.Split(line, ",")
			// Set the value in the specified cell
			var littleNo = count % 96
			var rIdx = littleNo % 8
			var cIdx = littleNo / 8
			if cells[0] != "covering" {
				err = excel.SetCellStr("引物订购单", "D"+strconv.Itoa(17+count), cells[0])
				if err != nil {
					return err
				}
				err = excel.SetCellStr("引物订购单", "E"+strconv.Itoa(17+count), cells[1])
				if err != nil {
					return err
				}
				// fmt.Printf("\t%c%d:%c%d:%s\n", 'D'+cIdx, panelRow+rIdx+2, 'A'+rIdx, cIdx+1, cells[0])
				SetCellStrStyle(
					selfSyntheticExcel,
					panelSheetName,
					fmt.Sprintf("%s\n(%d)", cells[0], len(cells[1])),
					panelCol+cIdx+1,
					panelRow+rIdx+2,
					bgStyleCurrent,
				)
			}
			count++
		}

		lines, err = ReadFileToLineArray(seq[2])
		if err != nil {
			return err
		}
		_, err = fmt.Fprintln(TailChangingPrimers, strings.Join(lines, "\n"))
		if err != nil {
			return err
		}
	}
	for _, f := range fList {
		err := f.Close()
		if err != nil {
			return err
		}
	}

	err = excel.UpdateLinkedValue()
	if err != nil {
		return err
	}

	// Save the changes back to the file
	err = excel.Save()
	if err != nil {
		return err
	}

	err = selfSyntheticExcel.SaveAs(prefix + "-自合.xlsx")
	if err != nil {
		log.Printf("SaveAs [%s-自合.xlsx] fail:[%v]", prefix, err)
		return err
	}

	return nil
}

func SetCellStr(excel *excelize.File, sheetName, valueStr string, col, row int) (err error) {
	cellName, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return
	}
	err = excel.SetCellStr(sheetName, cellName, valueStr)
	if err != nil {
		return
	}
	return
}

func SetCellStrStyle(excel *excelize.File, sheetName, valueStr string, col, row, styleId int) (err error) {
	cellName, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return
	}
	err = excel.SetCellStr(sheetName, cellName, valueStr)
	if err != nil {
		return
	}
	err = excel.SetCellStyle(sheetName, cellName, cellName, styleId)
	if err != nil {
		return
	}

	return
}

func SetSheetRow(excel *excelize.File, sheetName string, col, row int, slice any) (err error) {
	cellName, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return
	}
	err = excel.SetSheetRow(sheetName, cellName, slice)
	if err != nil {
		return
	}
	return
}

func SetSheetCol(excel *excelize.File, sheetName string, col, row int, slice interface{}) (err error) {
	cellName, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return
	}
	err = excel.SetSheetCol(sheetName, cellName, slice)
	if err != nil {
		return
	}
	return
}

func CreatePanelTable(excel *excelize.File, sheetName, panelName string, col, row, styleId int) (err error) {
	// 第一行
	err = SetSheetRow(excel, sheetName, col, row, &[]string{"板号：", panelName})
	if err != nil {
		log.Printf("SetSheetRow [%s:%d,%d][%v]:[%v]", sheetName, col, row, []string{"板号：", panelName}, err)
		return
	}

	// 第二行
	err = SetSheetRow(excel, sheetName, col+1, row+1, &[]int{1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12})
	if err != nil {
		log.Printf("SetSheetRow [%s:%d,%d][%v]:[%v]", sheetName, col+1, row+1, []int{1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, err)
		return
	}

	// 第一列
	err = SetSheetCol(excel, sheetName, col, row+2, &[]string{"A", "B", "C", "D", "E", "F", "G", "H"})
	if err != nil {
		log.Printf("SetSheetCol [%s:%d,%d][%v]:[%v]", sheetName, col, row+2, []string{"A", "B", "C", "D", "E", "F", "G", "H"}, err)
		return
	}

	hCell, err := excelize.CoordinatesToCellName(col, row+1)
	if err != nil {
		return
	}
	vCell, err := excelize.CoordinatesToCellName(col+12, row+9)
	if err != nil {
		return
	}
	err = excel.SetCellStyle(sheetName, hCell, vCell, styleId)
	if err != nil {
		return
	}

	return
}

func ReadFileToLineArray(filename string) ([]string, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	scanner := bufio.NewScanner(file)
	var lines []string
	for scanner.Scan() {
		lines = append(lines, scanner.Text())
	}
	if err := scanner.Err(); err != nil {
		return nil, err
	}

	return lines, nil
}

func CopyFile(source string, destination string) error {
	// Open the source file
	srcFile, err := os.Open(source)
	if err != nil {
		return err
	}
	defer srcFile.Close()

	// Create the destination file
	dstFile, err := os.Create(destination)
	if err != nil {
		return err
	}
	defer dstFile.Close()

	// Copy the contents of the source file to the destination file
	_, err = io.Copy(dstFile, srcFile)
	if err != nil {
		return err
	}

	// Flush any buffered data to ensure the file is fully written
	err = dstFile.Sync()
	if err != nil {
		return err
	}

	return nil
}

func calculateMD5Hash(s string) string {
	hasher := md5.New()
	hasher.Write([]byte(s))
	hash := hasher.Sum(nil)
	return hex.EncodeToString(hash)
}
