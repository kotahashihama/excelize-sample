package main

import (
	"fmt"
	"log"
	"os"
	"strings"

	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"

	"github.com/joho/godotenv"
	"github.com/xuri/excelize/v2"
)

const sheet = "Sheet1" // Sheet1 しか使わないので定数化

func main() {
	err := godotenv.Load()
	if err != nil {
		log.Fatal(err)
	}

	// ここをコメントイン・コメントアウトしながら go run main.go して確認
	createTextOnlyButHugeExcelFile()
	// updateMasterExcelFile()
	// createExtractedExcelFile()
	// createMasterExcelFile()
	// createHelloWorldExcelFile()
}

func createTextOnlyButHugeExcelFile() {
	f := excelize.NewFile()
	desktopPath := os.Getenv("ABSOLUTE_PATH_TO_DESKTOP")

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
			return
		}
	}()

	// 見出し
	f.SetCellValue(sheet, "A1", "店舗 ID")
	f.SetCellValue(sheet, "B1", "店舗名")
	f.SetCellValue(sheet, "C1", "完了日時")
	f.SetCellValue(sheet, "D1", "更新日時")
	f.SetCellValue(sheet, "E1", "報告者")
	f.SetCellValue(sheet, "F1", "画像1")
	f.SetCellValue(sheet, "G1", "画像2")
	f.SetCellValue(sheet, "H1", "画像3")
	f.SetCellValue(sheet, "I1", "画像4")
	f.SetCellValue(sheet, "J1", "画像5")

	// 内容
	rowCount := 22000
	loopedChar := "a"
	colPrefixes := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"}
	for i := 2; i < rowCount; i++ {
		for _, c := range colPrefixes {
			f.SetCellValue(sheet, fmt.Sprintf("%s%d", c, i), strings.Repeat(loopedChar, 10000))
		}
	}

	if err := f.SaveAs(fmt.Sprintf("%s/text_only_but_huge.xlsx", desktopPath)); err != nil {
		fmt.Println(err)
		return
	}
}

func updateMasterExcelFile() {
	desktopPath := os.Getenv("ABSOLUTE_PATH_TO_DESKTOP")

	f, err := excelize.OpenFile(os.Getenv("ABSOLUTE_PATH_TO_DESKTOP") + "/master.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
			return
		}
	}()

	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println(err)
		return
	}

	// 行を特定する
	targetId := "ididid"
	var rowNum int
	for i, row := range rows {
		if row[0] == targetId {
			rowNum = i + 1 // NOTE: 1行目は見出しなので、+1 する
			break
		}
	}

	// 「テキスト1」カラムを書き換える bar -> yahoo
	// WARNING: カラム名はユーザー任意の値なので、実際はこれで特定するのは避ける
	f.SetCellValue(sheet, fmt.Sprintf("F%d", rowNum), "yahoo")

	if err := f.SaveAs(fmt.Sprintf("%s/master.xlsx", desktopPath)); err != nil {
		fmt.Println(err)
		return
	}
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
	f.SetCellValue(sheet, "A1", "店舗 ID")
	f.SetCellValue(sheet, "B1", "店舗名")
	f.SetCellValue(sheet, "C1", "完了日時")
	f.SetCellValue(sheet, "D1", "更新日時")
	f.SetCellValue(sheet, "E1", "報告者")
	f.SetCellValue(sheet, "F1", "テキスト1")
	f.SetCellValue(sheet, "G1", "テキスト2")
	f.SetCellValue(sheet, "H1", "画像1")
	f.SetCellValue(sheet, "I1", "画像2")
	f.SetCellValue(sheet, "J1", "画像3")

	// マスターファイルを開く
	masterFile, err := excelize.OpenFile(fmt.Sprintf("%s/master.xlsx", desktopPath))
	if err != nil {
		fmt.Println(err)
		return
	}

	// マスターファイルから画像を取得
	pics, err := masterFile.GetPictures(sheet, "H2")
	if err != nil {
		fmt.Println(err)
		return
	}

	// 抽出先ファイルへ画像を挿入
	if err := f.AddPictureFromBytes(sheet, "H2", &excelize.Picture{
		Extension: pics[0].Extension,
		File:      pics[0].File,
		Format: &excelize.GraphicOptions{
			AutoFit: true, // NOTE: 抽出元から引き継がれないので、あらためて指定
		},
	}); err != nil {
		fmt.Println(err)
		return
	}

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

	// 内容1行目
	f.SetCellValue(sheet, "A2", "dididi")
	f.SetCellValue(sheet, "B2", "hoge")
	f.SetCellValue(sheet, "C2", "fuga")
	f.SetCellValue(sheet, "D2", "piyo")
	f.SetCellValue(sheet, "E2", "foo")
	f.SetCellValue(sheet, "F2", "bar")
	f.SetCellValue(sheet, "G2", "baz")
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
	if err := f.AddPicture(sheet, "J2", fmt.Sprintf("%s/image.png", desktopPath), &excelize.GraphicOptions{
		AutoFit: true,
	}); err != nil {
		fmt.Println(err)
		return
	}

	// 内容2行目
	f.SetCellValue(sheet, "A3", "ididid")
	f.SetCellValue(sheet, "B3", "hoge")
	f.SetCellValue(sheet, "C3", "fuga")
	f.SetCellValue(sheet, "D3", "piyo")
	f.SetCellValue(sheet, "E3", "foo")
	f.SetCellValue(sheet, "F3", "bar")
	f.SetCellValue(sheet, "G3", "baz")
	if err := f.AddPicture(sheet, "H3", fmt.Sprintf("%s/image.png", desktopPath), &excelize.GraphicOptions{
		AutoFit: true,
	}); err != nil {
		fmt.Println(err)
		return
	}
	if err := f.AddPicture(sheet, "I3", fmt.Sprintf("%s/image.png", desktopPath), &excelize.GraphicOptions{
		AutoFit: true,
	}); err != nil {
		fmt.Println(err)
		return
	}
	if err := f.AddPicture(sheet, "J3", fmt.Sprintf("%s/image.png", desktopPath), &excelize.GraphicOptions{
		AutoFit: true,
	}); err != nil {
		fmt.Println(err)
		return
	}

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
