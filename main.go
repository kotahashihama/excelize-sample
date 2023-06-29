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

	createExtractedExcelFile()
	// createMasterExcelFile()
	// createHelloWorldExcelFile()
}

func createExtractedExcelFile() {
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

	// マスターファイルを開く
	masterFile, err := excelize.OpenFile(fmt.Sprintf("%s/master.xlsx", desktopPath))
	if err != nil {
		fmt.Println(err)
		return
	}

	// マスターファイルから画像を取得
	pics, err := masterFile.GetPictures(sheet, "G2")
	if err != nil {
		fmt.Println(err)
		return
	}

	// 抽出先ファイルへ画像を挿入
	if err := f.AddPictureFromBytes(sheet, "G2", &excelize.Picture{
		Extension: pics[0].Extension,
		File:      pics[0].File,
		Format: &excelize.GraphicOptions{
			AutoFit: true, // NOTE: 抽出元から引き継がれないので、あらためて指定
		},
	}); err != nil {
		fmt.Println(err)
		return
	}

	// 抽出先ファイルを保存
	if err := f.SaveAs(fmt.Sprintf("%s/extracted.xlsx", desktopPath)); err != nil {
		fmt.Println(err)
		return
	}
}

// マスターファイルを作成
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
	f.SetCellValue(sheet, "B2", "fuga")
	f.SetCellValue(sheet, "C2", "piyo")
	f.SetCellValue(sheet, "D2", "foo")
	f.SetCellValue(sheet, "E2", "bar")
	f.SetCellValue(sheet, "F2", "baz")
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

// 使い方を確認するためのサンプル
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
