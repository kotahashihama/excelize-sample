package main

import (
	"fmt"
	"log"
	"os"

	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"

	"github.com/joho/godotenv"
	"github.com/xuri/excelize/v2"
)

const sheet = "Sheet1"

func main() {
	err := godotenv.Load()
	if err != nil {
		log.Fatal(err)
	}

	createMasterExcelFile()
	// createHelloWorldExcelFile()
}

func createMasterExcelFile() {
	f := excelize.NewFile()
	desktopPath := os.Getenv("ABSOLUTE_PATH_TO_DESKTOP")

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
			return
		}
	}()

	// 見出し
	f.SetCellValue(sheet, "A1", "店舗名")
	f.SetCellValue(sheet, "B1", "完了日時")
	f.SetCellValue(sheet, "C1", "更新日時")
	f.SetCellValue(sheet, "D1", "報告者")
	f.SetCellValue(sheet, "E1", "テキスト1")
	f.SetCellValue(sheet, "F1", "テキスト2")
	f.SetCellValue(sheet, "G1", "画像1")
	f.SetCellValue(sheet, "H1", "画像2")
	f.SetCellValue(sheet, "I1", "画像3")

	// 内容
	f.SetCellValue(sheet, "A2", "hoge")
	f.SetCellValue(sheet, "B2", "hoge")
	f.SetCellValue(sheet, "C2", "hoge")
	f.SetCellValue(sheet, "D2", "hoge")
	f.SetCellValue(sheet, "E2", "hoge")
	f.SetCellValue(sheet, "F2", "hoge")
	if err := f.AddPicture(sheet, "G2", fmt.Sprintf("%s/image.png", desktopPath), &excelize.GraphicOptions{
		AutoFit: true,
	}); err != nil {
		fmt.Println(err)
		return
	}
	if err := f.AddPicture(sheet, "H2", fmt.Sprintf("%s/image.png", desktopPath), &excelize.GraphicOptions{
		AutoFit: true,
	}); err != nil {
		fmt.Println(err)
		return
	}
	if err := f.AddPicture(sheet, "I2", fmt.Sprintf("%s/image.png", desktopPath), &excelize.GraphicOptions{
		AutoFit: true,
	}); err != nil {
		fmt.Println(err)
		return
	}

	// ファイルを保存
	if err := f.SaveAs(fmt.Sprintf("%s/master.xlsx", desktopPath)); err != nil {
		fmt.Println(err)
		return
	}
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
