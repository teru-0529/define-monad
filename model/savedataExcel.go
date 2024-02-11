package model

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/samber/lo"
	"github.com/xuri/excelize/v2"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
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

// toExcel(sht-elements)
func (savedata *SaveData) ToExcelElements() error {
	for _, element := range savedata.Elements {
		element.toExcel()
	}

	return nil
}

func (element *Element) toExcel() {
	t := japanese.ShiftJIS.NewEncoder()
	ary := []string{
		element.NameJp,
		element.NameEn,
		string(element.Domain),
		lo.Ternary(lo.IsNil(element.RegEx), "", *element.RegEx),
		int2Str(element.MinDigits),
		int2Str(element.MaxDigits),
		int2Str(element.MinValue),
		int2Str(element.MaxValue),
		element.Example,
		element.Description,
	}
	sjis, _, _ := transform.String(t, strings.Join(ary, "\t"))
	fmt.Println(sjis)
}

// toExcel(sht-derive-elements)
func (savedata *SaveData) ToExcelDeriveElements() error {
	for _, element := range savedata.DeliveElements {
		element.toExcel()
	}

	return nil
}

func (element *DeliveElement) toExcel() {
	t := japanese.ShiftJIS.NewEncoder()
	ary := []string{
		element.ref.NameJp,
		element.NameJp,
		element.NameEn,
		element.Description,
	}
	sjis, _, _ := transform.String(t, strings.Join(ary, "\t"))
	fmt.Println(sjis)
}

// toExcel(sht-derive-elements)
func (savedata *SaveData) ToExcelSegments() error {
	for _, element := range savedata.Segments {
		element.toExcel()
	}

	return nil
}

func (element *Segment) toExcel() {
	t := japanese.ShiftJIS.NewEncoder()
	ary := []string{
		element.ref.NameJp,
		element.Value,
		element.Name,
		element.Description,
	}
	sjis, _, _ := transform.String(t, strings.Join(ary, "\t"))
	fmt.Println(sjis)
}
