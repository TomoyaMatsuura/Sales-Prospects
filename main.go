package main

import (
	"io"
	"log"
	"os"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

// 売上実績セル関数
func calculateRow(i int, add int) string {
	row := [7]int{4, 5, 6, 7, 8, 9, 10}
	res := strconv.Itoa(row[i] + add)
	return res
}

// 売上見込セル関数
func calculateRowProspect(i int, add int) string {
	row := [7]int{4, 5, 6, 7, 8, 9, 10}
	res := strconv.Itoa(row[i] + add + i)
	return res
}

// ログ出力を行う関数
func loggingSettings(filename string) {
	logFile, _ := os.OpenFile(filename, os.O_RDWR|os.O_CREATE|os.O_APPEND, 0666)
	multiLogFile := io.MultiWriter(os.Stdout, logFile)
	log.SetFlags(log.Ldate | log.Ltime)
	log.SetOutput(multiLogFile)
}

func main() {
	branch := [7]string{"MBR", "MMX", "MCL", "MAR", "MLA", "MPE", "MCO"}

	loggingSettings("ログ.log")
	log.Println("-----    Start...    -----")

	for i := 0; i < len(branch); i++ {
		//保存されたファイルを開く
		filename := branch[i] + ".xlsx"
		branchFile, err := excelize.OpenFile(filename)
		if err != nil {
			log.Panicln(err)
		}
		defer func() {
			if err := branchFile.Close(); err != nil {
				log.Println(err)
			}
		}()

		// 1回目の売上予想取得
		salesResult, err := branchFile.GetCellValue("Sheet1", "B3")
		if err != nil {
			log.Println(err)
			return
		}
		salesProspect, err := branchFile.GetCellValue("Sheet1", "C3")
		if err != nil {
			log.Println(err)
			return
		}
		qtyProspect, err := branchFile.GetCellValue("Sheet1", "D3")
		if err != nil {
			log.Println(err)
			return
		}

		// 2回目の売上予想取得
		salesResult2, err := branchFile.GetCellValue("Sheet1", "G3")
		if err != nil {
			log.Println(err)
		}
		salesProspect2, err := branchFile.GetCellValue("Sheet1", "H3")
		if err != nil {
			log.Println(err)
		}
		qtyProspect2, err := branchFile.GetCellValue("Sheet1", "I3")
		if err != nil {
			log.Println(err)
		}


		// 転記先ファイルを開く
		sumFile, err := excelize.OpenFile("Summary.xlsx")
		if err != nil {
			log.Println(err)
		}
		defer func() {
			if err := sumFile.Close(); err != nil {
				log.Println(err)
			}
		}()

		month := time.Now().Month().String()

		// MCL1の場合/1000が必要
		// var clp int = 1
		// if i != 2 {
		// 	clp = 1000
		// }

		switch month {

		case "April":
			sumFile.SetCellValue("変数", "D"+calculateRow(i, 0), salesResult)
			sumFile.SetCellValue("変数", "D"+calculateRowProspect(i, 9), salesProspect)
			sumFile.SetCellValue("変数", "D"+calculateRowProspect(i, 10), qtyProspect)
			sumFile.SetCellValue("変数", "D"+calculateRow(i, 26), salesResult2)
			sumFile.SetCellValue("変数", "D"+calculateRowProspect(i, 35), salesProspect2)
			sumFile.SetCellValue("変数", "D"+calculateRowProspect(i, 36), qtyProspect2)
		case "May":
			sumFile.SetCellValue("変数", "E"+calculateRow(i, 0), salesResult)
			sumFile.SetCellValue("変数", "E"+calculateRowProspect(i, 9), salesProspect)
			sumFile.SetCellValue("変数", "E"+calculateRowProspect(i, 10), qtyProspect)
			sumFile.SetCellValue("変数", "E"+calculateRow(i, 26), salesResult2)
			sumFile.SetCellValue("変数", "E"+calculateRowProspect(i, 35), salesProspect2)
			sumFile.SetCellValue("変数", "E"+calculateRowProspect(i, 36), qtyProspect2)
		case "June":
			sumFile.SetCellValue("変数", "F"+calculateRow(i, 0), salesResult)
			sumFile.SetCellValue("変数", "F"+calculateRowProspect(i, 9), salesProspect)
			sumFile.SetCellValue("変数", "F"+calculateRowProspect(i, 10), qtyProspect)
			sumFile.SetCellValue("変数", "F"+calculateRow(i, 26), salesResult2)
			sumFile.SetCellValue("変数", "F"+calculateRowProspect(i, 35), salesProspect2)
			sumFile.SetCellValue("変数", "F"+calculateRowProspect(i, 36), qtyProspect2)
		case "July":
			sumFile.SetCellValue("変数", "G"+calculateRow(i, 0), salesResult)
			sumFile.SetCellValue("変数", "G"+calculateRowProspect(i, 9), salesProspect)
			sumFile.SetCellValue("変数", "G"+calculateRowProspect(i, 10), qtyProspect)
			sumFile.SetCellValue("変数", "G"+calculateRow(i, 26), salesResult2)
			sumFile.SetCellValue("変数", "G"+calculateRowProspect(i, 35), salesProspect2)
			sumFile.SetCellValue("変数", "G"+calculateRowProspect(i, 36), qtyProspect2)
		case "August":
			sumFile.SetCellValue("変数", "H"+calculateRow(i, 0), salesResult)
			sumFile.SetCellValue("変数", "H"+calculateRowProspect(i, 9), salesProspect)
			sumFile.SetCellValue("変数", "H"+calculateRowProspect(i, 10), qtyProspect)
			sumFile.SetCellValue("変数", "H"+calculateRow(i, 26), salesResult2)
			sumFile.SetCellValue("変数", "H"+calculateRowProspect(i, 35), salesProspect2)
			sumFile.SetCellValue("変数", "H"+calculateRowProspect(i, 36), qtyProspect2)
		case "September":
			sumFile.SetCellValue("変数", "I"+calculateRow(i, 0), salesResult)
			sumFile.SetCellValue("変数", "I"+calculateRowProspect(i, 9), salesProspect)
			sumFile.SetCellValue("変数", "I"+calculateRowProspect(i, 10), qtyProspect)
			sumFile.SetCellValue("変数", "I"+calculateRow(i, 26), salesResult2)
			sumFile.SetCellValue("変数", "I"+calculateRowProspect(i, 35), salesProspect2)
			sumFile.SetCellValue("変数", "I"+calculateRowProspect(i, 36), qtyProspect2)
		case "October":
			sumFile.SetCellValue("変数", "J"+calculateRow(i, 0), salesResult)
			sumFile.SetCellValue("変数", "J"+calculateRowProspect(i, 9), salesProspect)
			sumFile.SetCellValue("変数", "J"+calculateRowProspect(i, 10), qtyProspect)
			sumFile.SetCellValue("変数", "J"+calculateRow(i, 26), salesResult2)
			sumFile.SetCellValue("変数", "J"+calculateRowProspect(i, 35), salesProspect2)
			sumFile.SetCellValue("変数", "J"+calculateRowProspect(i, 36), qtyProspect2)
		case "November":
			sumFile.SetCellValue("変数", "K"+calculateRow(i, 0), salesResult)
			sumFile.SetCellValue("変数", "K"+calculateRowProspect(i, 9), salesProspect)
			sumFile.SetCellValue("変数", "K"+calculateRowProspect(i, 10), qtyProspect)
			sumFile.SetCellValue("変数", "K"+calculateRow(i, 26), salesResult2)
			sumFile.SetCellValue("変数", "K"+calculateRowProspect(i, 35), salesProspect2)
			sumFile.SetCellValue("変数", "K"+calculateRowProspect(i, 36), qtyProspect2)
		case "December":
			sumFile.SetCellValue("変数", "L"+calculateRow(i, 0), salesResult)
			sumFile.SetCellValue("変数", "L"+calculateRowProspect(i, 9), salesProspect)
			sumFile.SetCellValue("変数", "L"+calculateRowProspect(i, 10), qtyProspect)
			sumFile.SetCellValue("変数", "L"+calculateRow(i, 26), salesResult2)
			sumFile.SetCellValue("変数", "L"+calculateRowProspect(i, 35), salesProspect2)
			sumFile.SetCellValue("変数", "L"+calculateRowProspect(i, 36), qtyProspect2)
		case "January":
			sumFile.SetCellValue("変数", "M"+calculateRow(i, 0), salesResult)
			sumFile.SetCellValue("変数", "M"+calculateRowProspect(i, 9), salesProspect)
			sumFile.SetCellValue("変数", "M"+calculateRowProspect(i, 10), qtyProspect)
			sumFile.SetCellValue("変数", "M"+calculateRow(i, 26), salesResult2)
			sumFile.SetCellValue("変数", "M"+calculateRowProspect(i, 35), salesProspect2)
			sumFile.SetCellValue("変数", "M"+calculateRowProspect(i, 36), qtyProspect2)
		case "February":
			sumFile.SetCellValue("変数", "N"+calculateRow(i, 0), salesResult)
			sumFile.SetCellValue("変数", "N"+calculateRowProspect(i, 9), salesProspect)
			sumFile.SetCellValue("変数", "N"+calculateRowProspect(i, 10), qtyProspect)
			sumFile.SetCellValue("変数", "N"+calculateRow(i, 26), salesResult2)
			sumFile.SetCellValue("変数", "N"+calculateRowProspect(i, 35), salesProspect2)
			sumFile.SetCellValue("変数", "N"+calculateRowProspect(i, 36), qtyProspect2)
		case "March":
			sumFile.SetCellValue("変数", "O"+calculateRow(i, 0), salesResult)
			sumFile.SetCellValue("変数", "O"+calculateRowProspect(i, 9), salesProspect)
			sumFile.SetCellValue("変数", "O"+calculateRowProspect(i, 10), qtyProspect)
			sumFile.SetCellValue("変数", "O"+calculateRow(i, 26), salesResult2)
			sumFile.SetCellValue("変数", "O"+calculateRowProspect(i, 35), salesProspect2)
			sumFile.SetCellValue("変数", "O"+calculateRowProspect(i, 36), qtyProspect2)
		}

		if err := sumFile.Save(); err != nil {
			log.Println(err)
		}
		log.Println("----- "+ branch[i] + " Registered -----")
		

	}
	log.Println("----- ALL Completed! -----")
}
