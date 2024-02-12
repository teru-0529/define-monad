package elements

import (
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

const SHT_ELEMENTS = "項目"
const SHT_DERIVE_ELEMENTS = "別名"
const SHT_SEGMENTS = "区分値"

func FromExcel(fileName string) (*SaveData, error) {

	savedata := SaveData{}

	// INFO: ファイルの読み込み
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	// INFO: [項目]シートデータ取得
	rows, err := f.GetRows(SHT_ELEMENTS)
	if err != nil {
		return nil, err
	}
	for _, row := range rows[1:] { // ヘッダー行は読み飛ばす
		if row[0] == "" {
			// 空行は読み飛ばす
			continue
		}
		element := Element{}
		element.NameJp = row[0]
		element.NameEn = row[1]
		element.Domain = Dom(row[3])
		if row[4] != "" {
			element.RegEx = &row[4]
		}
		if row[5] != "" {
			str := strings.Replace(row[5], ",", "", -1)
			num, _ := strconv.Atoi(str)
			element.MinDigits = &num
		}
		if row[6] != "" {
			str := strings.Replace(row[6], ",", "", -1)
			num, _ := strconv.Atoi(str)
			element.MaxDigits = &num
		}
		if row[7] != "" {
			str := strings.Replace(row[7], ",", "", -1)
			num, _ := strconv.Atoi(str)
			element.MinValue = &num
		}
		if row[8] != "" {
			str := strings.Replace(row[8], ",", "", -1)
			num, _ := strconv.Atoi(str)
			element.MaxValue = &num
		}
		element.Example = row[9]
		element.Description = row[10]

		savedata.Elements = append(savedata.Elements, element)
	}

	// INFO: [別名]シートデータ取得
	rows, err = f.GetRows(SHT_DERIVE_ELEMENTS)
	if err != nil {
		return nil, err
	}
	for _, row := range rows[1:] { // ヘッダー行は読み飛ばす
		if row[0] == "" {
			// 空行は読み飛ばす
			continue
		}
		element := DeliveElement{}
		element.Origin = row[0]
		element.NameJp = row[1]
		element.NameEn = row[2]
		element.Description = row[4]

		savedata.DeliveElements = append(savedata.DeliveElements, element)
	}

	// INFO: [区分値]シートデータ取得
	rows, err = f.GetRows(SHT_SEGMENTS)
	if err != nil {
		return nil, err
	}
	for _, row := range rows[1:] { // ヘッダー行は読み飛ばす
		if row[0] == "" {
			// 空行は読み飛ばす
			continue
		}
		element := Segment{}
		element.Key = row[0]
		element.Value = row[1]
		element.Name = row[2]
		element.Description = row[3]

		savedata.Segments = append(savedata.Segments, element)
	}
	return &savedata, nil
}
