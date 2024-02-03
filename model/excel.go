package model

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

const shtNameElements = "項目"

func FromExcel(fileName string) (*SaveData, error) {

	// var savedata *SaveData
	savedata := SaveData{}

	// INFO: ファイルの読み込み
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	// INFO: データ取得
	rows, err := f.GetRows(shtNameElements)
	if err != nil {
		return nil, err
	}
	for _, row := range rows[1:] { // ヘッダー行は読み飛ばす
		if row[0] == "" {
			// 空行は読み飛ばす
			continue
		}
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
		}
		fmt.Println()

	}
	return &savedata, nil
}

// 項目オブジェクトの追加
// func excel2Elements(nameJp string) ([]Element, error) {
// 	for i := range savedata.Elements {
// 		if savedata.Elements[i].NameJp == nameJp {
// 			return &savedata.Elements[i], nil
// 		}
// 	}
// 	return nil, nil
// }
