package main

import (
	"fmt"
	"log"
	"os"

	"github.com/joho/godotenv"
	"github.com/xuri/excelize/v2"
)

func main() {
	err := godotenv.Load()
	if err != nil {
		log.Fatal("Error loading .env file")
	}

	createHelloWorldExcelFile()
}

func createHelloWorldExcelFile() {
	f := excelize.NewFile()

	// 関数終了の直前にクローズする
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// シートを作成
	index, err := f.NewSheet("Sheet2")
	if err != nil {
		fmt.Println(err)
		return
	}

	// セルに値をセット
	f.SetCellValue("Sheet2", "A2", "Hello world.")
	f.SetCellValue("Sheet1", "B2", 100)

	// アクティブなシートを切り替える
	f.SetActiveSheet(index)

	// ファイルを保存
	desktopPath := os.Getenv("ABSOLUTE_PATH_TO_DESKTOP")
	if err := f.SaveAs(fmt.Sprintf("%s/Book1.xlsx", desktopPath)); err != nil {
		fmt.Println(err)
	}
}
