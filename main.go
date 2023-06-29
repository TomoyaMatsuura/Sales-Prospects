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

// コメント出力関数
func setComment(sumFile *excelize.File, commentNumber int, commentValue [10]string, sheetName string) {
	for m := 0; m < 10; m++ {
		sumFile.SetCellValue(sheetName, calculateComment(commentNumber+m), commentValue[m])
	}
}

func main() {
	branch := [7]string{"MBR", "MMX", "MCL", "MAR", "MLA", "MPE", "MCO"}
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
		// salesReport, err := branchFile.GetCellValue("速報値", "B6")
		// if err != nil {
		// 	log.Println(err)
		// }
		// qtyReport, err := branchFile.GetCellValue("速報値", "D6")
		// if err != nil {
		// 	log.Println(err)
		// }
		// commentReport := [10]string{}
		// for l := 0; l < 10; l++ {
		// 	comment, err := branchFile.GetCellValue("速報値", comment[l])
		// 	if err != nil {
		// 		log.Println(err)
		// 		return
		// 	}
		// 	commentReport[l] = comment
		// }
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
			sumFile.SetCellValue("変数", "D"+calculateResult(rowNumber, 0), salesResult)
			sumFile.SetCellValue("変数", "D"+calculateProspect(rowNumber, 9, addNumber), salesProspect)
			sumFile.SetCellValue("変数", "D"+calculateProspect(rowNumber, 10, addNumber), qtyProspect)
			sumFile.SetCellValue("変数", "D"+calculateResult(rowNumber, 26), salesResult2)
			sumFile.SetCellValue("変数", "D"+calculateProspect(rowNumber, 35, addNumber), salesProspect2)
			sumFile.SetCellValue("変数", "D"+calculateProspect(rowNumber, 36, addNumber), qtyProspect2)
			sumFile.SetCellValue("変数", "O"+calculateProspect(rowNumber, 52, addNumber), salesReport)
			sumFile.SetCellValue("変数", "O"+calculateProspect(rowNumber, 53, addNumber), qtyReport)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "March Report")

		case "May":
			sumFile.SetCellValue("変数", "E"+calculateResult(rowNumber, 0), salesResult)
			sumFile.SetCellValue("変数", "E"+calculateProspect(rowNumber, 9, addNumber), salesProspect)
			sumFile.SetCellValue("変数", "E"+calculateProspect(rowNumber, 10, addNumber), qtyProspect)
			sumFile.SetCellValue("変数", "E"+calculateResult(rowNumber, 26), salesResult2)
			sumFile.SetCellValue("変数", "E"+calculateProspect(rowNumber, 35, addNumber), salesProspect2)
			sumFile.SetCellValue("変数", "E"+calculateProspect(rowNumber, 36, addNumber), qtyProspect2)
			sumFile.SetCellValue("変数", "D"+calculateProspect(rowNumber, 52, addNumber), salesReport)
			sumFile.SetCellValue("変数", "D"+calculateProspect(rowNumber, 53, addNumber), qtyReport)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "April Report")
		case "June":
			sumFile.SetCellValue("変数", "F"+calculateResult(rowNumber, 0), salesResult)
			sumFile.SetCellValue("変数", "F"+calculateProspect(rowNumber, 9, addNumber), salesProspect)
			sumFile.SetCellValue("変数", "F"+calculateProspect(rowNumber, 10, addNumber), qtyProspect)
			sumFile.SetCellValue("変数", "F"+calculateResult(rowNumber, 26), salesResult2)
			sumFile.SetCellValue("変数", "F"+calculateProspect(rowNumber, 35, addNumber), salesProspect2)
			sumFile.SetCellValue("変数", "F"+calculateProspect(rowNumber, 36, addNumber), qtyProspect2)
			sumFile.SetCellValue("変数", "E"+calculateProspect(rowNumber, 52, addNumber), salesReport)
			sumFile.SetCellValue("変数", "E"+calculateProspect(rowNumber, 53, addNumber), qtyReport)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "May Report")

		case "July":
			sumFile.SetCellValue("変数", "G"+calculateResult(rowNumber, 0), salesResult)
			sumFile.SetCellValue("変数", "G"+calculateProspect(rowNumber, 9, addNumber), salesProspect)
			sumFile.SetCellValue("変数", "G"+calculateProspect(rowNumber, 10, addNumber), qtyProspect)
			sumFile.SetCellValue("変数", "G"+calculateResult(rowNumber, 26), salesResult2)
			sumFile.SetCellValue("変数", "G"+calculateProspect(rowNumber, 35, addNumber), salesProspect2)
			sumFile.SetCellValue("変数", "G"+calculateProspect(rowNumber, 36, addNumber), qtyProspect2)
			sumFile.SetCellValue("変数", "F"+calculateProspect(rowNumber, 52, addNumber), salesReport)
			sumFile.SetCellValue("変数", "F"+calculateProspect(rowNumber, 53, addNumber), qtyReport)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "June Report")
		case "August":
			sumFile.SetCellValue("変数", "H"+calculateResult(rowNumber, 0), salesResult)
			sumFile.SetCellValue("変数", "H"+calculateProspect(rowNumber, 9, addNumber), salesProspect)
			sumFile.SetCellValue("変数", "H"+calculateProspect(rowNumber, 10, addNumber), qtyProspect)
			sumFile.SetCellValue("変数", "H"+calculateResult(rowNumber, 26), salesResult2)
			sumFile.SetCellValue("変数", "H"+calculateProspect(rowNumber, 35, addNumber), salesProspect2)
			sumFile.SetCellValue("変数", "H"+calculateProspect(rowNumber, 36, addNumber), qtyProspect2)
			sumFile.SetCellValue("変数", "G"+calculateProspect(rowNumber, 52, addNumber), salesReport)
			sumFile.SetCellValue("変数", "G"+calculateProspect(rowNumber, 53, addNumber), qtyReport)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "July Report")
		case "September":
			sumFile.SetCellValue("変数", "I"+calculateResult(rowNumber, 0), salesResult)
			sumFile.SetCellValue("変数", "I"+calculateProspect(rowNumber, 9, addNumber), salesProspect)
			sumFile.SetCellValue("変数", "I"+calculateProspect(rowNumber, 10, addNumber), qtyProspect)
			sumFile.SetCellValue("変数", "I"+calculateResult(rowNumber, 26), salesResult2)
			sumFile.SetCellValue("変数", "I"+calculateProspect(rowNumber, 35, addNumber), salesProspect2)
			sumFile.SetCellValue("変数", "I"+calculateProspect(rowNumber, 36, addNumber), qtyProspect2)
			sumFile.SetCellValue("変数", "H"+calculateProspect(rowNumber, 52, addNumber), salesReport)
			sumFile.SetCellValue("変数", "H"+calculateProspect(rowNumber, 53, addNumber), qtyReport)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "August Report")
		case "October":
			sumFile.SetCellValue("変数", "J"+calculateResult(rowNumber, 0), salesResult)
			sumFile.SetCellValue("変数", "J"+calculateProspect(rowNumber, 9, addNumber), salesProspect)
			sumFile.SetCellValue("変数", "J"+calculateProspect(rowNumber, 10, addNumber), qtyProspect)
			sumFile.SetCellValue("変数", "J"+calculateResult(rowNumber, 26), salesResult2)
			sumFile.SetCellValue("変数", "J"+calculateProspect(rowNumber, 35, addNumber), salesProspect2)
			sumFile.SetCellValue("変数", "J"+calculateProspect(rowNumber, 36, addNumber), qtyProspect2)
			sumFile.SetCellValue("変数", "I"+calculateProspect(rowNumber, 52, addNumber), salesReport)
			sumFile.SetCellValue("変数", "I"+calculateProspect(rowNumber, 53, addNumber), qtyReport)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "September Report")
		case "November":
			sumFile.SetCellValue("変数", "K"+calculateResult(rowNumber, 0), salesResult)
			sumFile.SetCellValue("変数", "K"+calculateProspect(rowNumber, 9, addNumber), salesProspect)
			sumFile.SetCellValue("変数", "K"+calculateProspect(rowNumber, 10, addNumber), qtyProspect)
			sumFile.SetCellValue("変数", "K"+calculateResult(rowNumber, 26), salesResult2)
			sumFile.SetCellValue("変数", "K"+calculateProspect(rowNumber, 35, addNumber), salesProspect2)
			sumFile.SetCellValue("変数", "K"+calculateProspect(rowNumber, 36, addNumber), qtyProspect2)
			sumFile.SetCellValue("変数", "J"+calculateProspect(rowNumber, 52, addNumber), salesReport)
			sumFile.SetCellValue("変数", "J"+calculateProspect(rowNumber, 53, addNumber), qtyReport)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "October Report")
		case "December":
			sumFile.SetCellValue("変数", "L"+calculateResult(rowNumber, 0), salesResult)
			sumFile.SetCellValue("変数", "L"+calculateProspect(rowNumber, 9, addNumber), salesProspect)
			sumFile.SetCellValue("変数", "L"+calculateProspect(rowNumber, 10, addNumber), qtyProspect)
			sumFile.SetCellValue("変数", "L"+calculateResult(rowNumber, 26), salesResult2)
			sumFile.SetCellValue("変数", "L"+calculateProspect(rowNumber, 35, addNumber), salesProspect2)
			sumFile.SetCellValue("変数", "L"+calculateProspect(rowNumber, 36, addNumber), qtyProspect2)
			sumFile.SetCellValue("変数", "K"+calculateProspect(rowNumber, 52, addNumber), salesReport)
			sumFile.SetCellValue("変数", "K"+calculateProspect(rowNumber, 53, addNumber), qtyReport)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "November Report")
		case "January":
			sumFile.SetCellValue("変数", "M"+calculateResult(rowNumber, 0), salesResult)
			sumFile.SetCellValue("変数", "M"+calculateProspect(rowNumber, 9, addNumber), salesProspect)
			sumFile.SetCellValue("変数", "M"+calculateProspect(rowNumber, 10, addNumber), qtyProspect)
			sumFile.SetCellValue("変数", "M"+calculateResult(rowNumber, 26), salesResult2)
			sumFile.SetCellValue("変数", "M"+calculateProspect(rowNumber, 35, addNumber), salesProspect2)
			sumFile.SetCellValue("変数", "M"+calculateProspect(rowNumber, 36, addNumber), qtyProspect2)
			sumFile.SetCellValue("変数", "L"+calculateProspect(rowNumber, 52, addNumber), salesReport)
			sumFile.SetCellValue("変数", "L"+calculateProspect(rowNumber, 53, addNumber), qtyReport)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "December Report")
		case "February":
			sumFile.SetCellValue("変数", "N"+calculateResult(rowNumber, 0), salesResult)
			sumFile.SetCellValue("変数", "N"+calculateProspect(rowNumber, 9, addNumber), salesProspect)
			sumFile.SetCellValue("変数", "N"+calculateProspect(rowNumber, 10, addNumber), qtyProspect)
			sumFile.SetCellValue("変数", "N"+calculateResult(rowNumber, 26), salesResult2)
			sumFile.SetCellValue("変数", "N"+calculateProspect(rowNumber, 35, addNumber), salesProspect2)
			sumFile.SetCellValue("変数", "N"+calculateProspect(rowNumber, 36, addNumber), qtyProspect2)
			sumFile.SetCellValue("変数", "M"+calculateProspect(rowNumber, 52, addNumber), salesReport)
			sumFile.SetCellValue("変数", "M"+calculateProspect(rowNumber, 53, addNumber), qtyReport)
			setComment(sumFile, commentNumber, commentProspect, month+" 1st")
			setComment(sumFile, commentNumber, commentProspect2, month+" 2nd")
			setComment(sumFile, commentNumber, commentReport, "January Report")
		case "March":
			sumFile.SetCellValue("変数", "O"+calculateResult(rowNumber, 0), salesResult)
			sumFile.SetCellValue("変数", "O"+calculateProspect(rowNumber, 9, addNumber), salesProspect)
			sumFile.SetCellValue("変数", "O"+calculateProspect(rowNumber, 10, addNumber), qtyProspect)
			sumFile.SetCellValue("変数", "O"+calculateResult(rowNumber, 26), salesResult2)
			sumFile.SetCellValue("変数", "O"+calculateProspect(rowNumber, 35, addNumber), salesProspect2)
			sumFile.SetCellValue("変数", "O"+calculateProspect(rowNumber, 36, addNumber), qtyProspect2)
			sumFile.SetCellValue("変数", "N"+calculateProspect(rowNumber, 52, addNumber), salesReport)
			sumFile.SetCellValue("変数", "N"+calculateProspect(rowNumber, 53, addNumber), qtyReport)
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
