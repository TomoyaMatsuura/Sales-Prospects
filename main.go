package main

import (
	"io"
	"log"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// 売上実績セル関数
func calculateResult(i int, add int) string {
	res := strconv.Itoa(i + add)
	return res
}

// 売上予想と売上速報セル関数
func calculateProspect(i int, add int, addNumber int) string {
	res := strconv.Itoa(i + add + addNumber)
	return res
}

// コメントセル関数
func calculateComment(i int) string {
	res := strconv.Itoa(i)
	return "A" + res
}

// ログ出力を行う関数
func loggingSettings(filename string) {
	logFile, _ := os.OpenFile(filename, os.O_RDWR|os.O_CREATE|os.O_APPEND, 0666)
	multiLogFile := io.MultiWriter(os.Stdout, logFile)
	log.SetFlags(log.Ldate | log.Ltime)
	log.SetOutput(multiLogFile)
}

// m秒待機する関数
func sleep(m int) {
	time.Sleep(time.Duration(m) * time.Second)
}

// 売上予想セル取得
func getCellProspect(branchFile *excelize.File, comment [10]string, sheet string) (string, string, string, [10]string) {
	salesResult, err := branchFile.GetCellValue(sheet, "B5")
	if err != nil {
		log.Println(err)
	}
	salesProspect, err := branchFile.GetCellValue(sheet, "B10")
	if err != nil {
		log.Println(err)
	}
	qtyProspect, err := branchFile.GetCellValue(sheet, "D10")
	if err != nil {
		log.Println(err)
	}
	commentProspect := [10]string{}
	for j := 0; j < 10; j++ {
		comment1, err := branchFile.GetCellValue(sheet, comment[j])
		if err != nil {
			log.Println(err)
		}
		commentProspect[j] = comment1
	}
	return salesResult, salesProspect, qtyProspect, commentProspect
}

// 売上速報セル取得
func getCellReport(branchFile *excelize.File, comment [10]string, sheet string) (string, string, [10]string) {
	salesReport, err := branchFile.GetCellValue(sheet, "B6")
	if err != nil {
		log.Println(err)
	}
	qtyReport, err := branchFile.GetCellValue(sheet, "D6")
	if err != nil {
		log.Println(err)
	}
	commentReport := [10]string{}
	for j := 0; j < 10; j++ {
		comment1, err := branchFile.GetCellValue(sheet, comment[j])
		if err != nil {
			log.Println(err)
		}
		commentReport[j] = comment1
	}
	return salesReport, qtyReport, commentReport
}

// 売上予想出力関数
func setProspect(sumFile *excelize.File, column string, rowNumber int, salesResult string, salesProspect string, qtyProspect string, addNumber int) {
	if salesResult != "" {
		sumFile.SetCellValue("変数", column+calculateResult(rowNumber, 0), salesResult)
	}
	if salesProspect != "" {
		sumFile.SetCellValue("変数", column+calculateProspect(rowNumber, 9, addNumber), salesProspect)
	}

	if qtyProspect != "" {
		sumFile.SetCellValue("変数", column+calculateProspect(rowNumber, 10, addNumber), qtyProspect)
	}
}

func setProspect2(sumFile *excelize.File, column string, rowNumber int, salesResult2 string, salesProspect2 string, qtyProspect2 string, addNumber int) {
	if salesResult2 != "" {
		sumFile.SetCellValue("変数", column+calculateResult(rowNumber, 26), salesResult2)
	}

	if salesProspect2 != "" {
		sumFile.SetCellValue("変数", column+calculateProspect(rowNumber, 35, addNumber), salesProspect2)
	}

	if qtyProspect2 != "" {
		sumFile.SetCellValue("変数", column+calculateProspect(rowNumber, 36, addNumber), qtyProspect2)
	}

}

func setReport(sumFile *excelize.File, column string, rowNumber int, salesReport string, qtyReport string, addNumber int) {
	if salesReport != "" {
		sumFile.SetCellValue("変数", column+calculateProspect(rowNumber, 52, addNumber), salesReport)
	}

	if qtyReport != "" {
		sumFile.SetCellValue("変数", column+calculateProspect(rowNumber, 53, addNumber), qtyReport)
	}

}

// コメント出力関数
func setComment(sumFile *excelize.File, commentNumber int, commentValue [10]string, sheetName string) {
	for m := 0; m < 10; m++ {
		if commentValue[m] != "" {
			sumFile.SetCellValue(sheetName, calculateComment(commentNumber+m), commentValue[m])
		}
	}
}

func main() {
	branch := [7]string{"MBR", "MMX", "MCL", "MAR", "MLA", "MPE", "MCO"}

	// 予想/速報のコメントはA18->A27なので、関数で取得する用の配列
	comment := [10]string{"A18", "A19", "A20", "A21", "A22", "A23", "A24", "A25", "A26", "A27"}

	loggingSettings("ログ.log")
	log.Println("Start... ")

	//カレントディレクトリのファイル一覧を取得
	files, _ := os.ReadDir("./")

	//ディレクトリ内のファイルが販社名が含まれるか確認
	for i := 0; i < len(files); i++ {
		fileName := files[i].Name()

		//ファイル名とマッチするかどうか判定
		rowNumber := 0
		addNumber := 0
		commentNumber := 0

		for _, s := range branch {
			if strings.Contains(fileName, s) {
				switch s {
				case "MBR":
					rowNumber = 4
					addNumber = 0
					commentNumber = 23
				case "MMX":
					rowNumber = 5
					addNumber = 1
					commentNumber = 35
				case "MCL":
					rowNumber = 6
					addNumber = 2
					commentNumber = 47
				case "MAR":
					rowNumber = 7
					addNumber = 3
					commentNumber = 59
				case "MLA":
					rowNumber = 8
					addNumber = 4
					commentNumber = 71
				case "MPE":
					rowNumber = 9
					addNumber = 5
					commentNumber = 83
				case "MCO":
					rowNumber = 10
					addNumber = 6
					commentNumber = 95
				default:
					rowNumber = 0
				}
			}
		}
		if rowNumber == 0 {
			continue
		}

		//保存されたファイルを開く
		branchFile, err := excelize.OpenFile(fileName)
		if err != nil {
			log.Panicln(err)
		}
		defer func() {
			if err := branchFile.Close(); err != nil {
				log.Println(err)
			}
		}()

		// 1回目の売上予想取得
		salesResult, salesProspect, qtyProspect, commentProspect := getCellProspect(branchFile, comment, "売上予想1回目")
		// 2回目の売上予想取得
		salesResult2, salesProspect2, qtyProspect2, commentProspect2 := getCellProspect(branchFile, comment, "売上予想2回目")
		// 売上速報取得
		salesReport, qtyReport, commentReport := getCellReport(branchFile, comment, "速報値")
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
			setProspect(sumFile, "D", rowNumber, salesResult, salesProspect, qtyProspect, addNumber)
			setProspect2(sumFile, "D", rowNumber, salesResult2, salesProspect2, qtyProspect2, addNumber)
			setReport(sumFile, "O", rowNumber, salesReport, qtyReport, addNumber)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "March Report")

		case "May":
			setProspect(sumFile, "E", rowNumber, salesResult, salesProspect, qtyProspect, addNumber)
			setProspect2(sumFile, "E", rowNumber, salesResult2, salesProspect2, qtyProspect2, addNumber)
			setReport(sumFile, "D", rowNumber, salesReport, qtyReport, addNumber)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "April Report")
		case "June":
			setProspect(sumFile, "F", rowNumber, salesResult, salesProspect, qtyProspect, addNumber)
			setProspect2(sumFile, "F", rowNumber, salesResult2, salesProspect2, qtyProspect2, addNumber)
			setReport(sumFile, "E", rowNumber, salesReport, qtyReport, addNumber)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "May Report")

		case "July":
			setProspect(sumFile, "G", rowNumber, salesResult, salesProspect, qtyProspect, addNumber)
			setProspect2(sumFile, "G", rowNumber, salesResult2, salesProspect2, qtyProspect2, addNumber)
			setReport(sumFile, "F", rowNumber, salesReport, qtyReport, addNumber)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "June Report")
		case "August":
			setProspect(sumFile, "H", rowNumber, salesResult, salesProspect, qtyProspect, addNumber)
			setProspect2(sumFile, "H", rowNumber, salesResult2, salesProspect2, qtyProspect2, addNumber)
			setReport(sumFile, "G", rowNumber, salesReport, qtyReport, addNumber)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "July Report")
		case "September":
			setProspect(sumFile, "I", rowNumber, salesResult, salesProspect, qtyProspect, addNumber)
			setProspect2(sumFile, "I", rowNumber, salesResult2, salesProspect2, qtyProspect2, addNumber)
			setReport(sumFile, "H", rowNumber, salesReport, qtyReport, addNumber)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "August Report")
		case "October":
			setProspect(sumFile, "J", rowNumber, salesResult, salesProspect, qtyProspect, addNumber)
			setProspect2(sumFile, "J", rowNumber, salesResult2, salesProspect2, qtyProspect2, addNumber)
			setReport(sumFile, "I", rowNumber, salesReport, qtyReport, addNumber)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "September Report")
		case "November":
			setProspect(sumFile, "K", rowNumber, salesResult, salesProspect, qtyProspect, addNumber)
			setProspect2(sumFile, "K", rowNumber, salesResult2, salesProspect2, qtyProspect2, addNumber)
			setReport(sumFile, "J", rowNumber, salesReport, qtyReport, addNumber)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "October Report")
		case "December":
			setProspect(sumFile, "L", rowNumber, salesResult, salesProspect, qtyProspect, addNumber)
			setProspect2(sumFile, "L", rowNumber, salesResult2, salesProspect2, qtyProspect2, addNumber)
			setReport(sumFile, "K", rowNumber, salesReport, qtyReport, addNumber)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "November Report")
		case "January":
			setProspect(sumFile, "M", rowNumber, salesResult, salesProspect, qtyProspect, addNumber)
			setProspect2(sumFile, "M", rowNumber, salesResult2, salesProspect2, qtyProspect2, addNumber)
			setReport(sumFile, "L", rowNumber, salesReport, qtyReport, addNumber)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "December Report")
		case "February":
			setProspect(sumFile, "N", rowNumber, salesResult, salesProspect, qtyProspect, addNumber)
			setProspect2(sumFile, "N", rowNumber, salesResult2, salesProspect2, qtyProspect2, addNumber)
			setReport(sumFile, "M", rowNumber, salesReport, qtyReport, addNumber)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "January Report")
		case "March":
			setProspect(sumFile, "O", rowNumber, salesResult, salesProspect, qtyProspect, addNumber)
			setProspect2(sumFile, "O", rowNumber, salesResult2, salesProspect2, qtyProspect2, addNumber)
			setReport(sumFile, "N", rowNumber, salesReport, qtyReport, addNumber)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "February Report")
		}

		if err := sumFile.Save(); err != nil {
			log.Println(err)
		}
		log.Println("----- " + fileName + " Registered -----")
	}
	log.Println("ALL Completed!")
	sleep(2)
}
