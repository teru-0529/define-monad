package model

import (
	"fmt"
	"strconv"

	"github.com/samber/lo"
	"github.com/teru-0529/define-monad/store"
)

// elements-viewの書き込み
func (savedata *SaveData) WriteViewElements(path string) error {
	// INFO: Writerの取得
	writer, cleanup, err := store.NewCsvWriter(path)
	if err != nil {
		return err
	}
	defer cleanup()

	// INFO: 書き込み
	defer writer.Flush() //内部バッファのフラッシュは必須
	writer.Write([]string{
		"名称(JP)",
		"名称(EN)",
		"ドメイン",
		"正規表現",
		"最小桁数",
		"最大桁数",
		"最小値",
		"最大値",
		"API(example)",
		"説明",
	})
	for _, elements := range savedata.Elements {
		if err := writer.Write(elements.toArray()); err != nil {
			return fmt.Errorf("cannot write record: %s", err.Error())
		}
	}

	return nil
}

func (element *Element) toArray() []string {
	return []string{
		element.NameJp,
		element.NameEn,
		string(element.Domain),
		lo.Ternary(lo.IsNil(element.RegEx), "", *element.RegEx),
		int2Str(element.MinDigits),
		int2Str(element.MaxDigits),
		int2Str(element.MinValue),
		int2Str(element.Maxvalue),
		element.Example,
		element.Description,
	}
}

func int2Str(val *int) string {
	return lo.TernaryF(
		lo.IsNil(val),
		func() string { return "" },
		func() string { return strconv.Itoa(*val) },
	)
}

// derive-elements-viewの書き込み
func (savedata *SaveData) WriteViewDeriveElements(path string) error {
	// INFO: Writerの取得
	writer, cleanup, err := store.NewCsvWriter(path)
	if err != nil {
		return err
	}
	defer cleanup()

	// INFO: 書き込み
	defer writer.Flush() //内部バッファのフラッシュは必須
	writer.Write([]string{
		"派生元項目(JP)",
		"派生元項目(EN)",
		"名称(JP)",
		"名称(EN)",
		"説明",
	})
	for _, elements := range savedata.DeliveElements {
		if err := writer.Write(elements.toArray()); err != nil {
			return fmt.Errorf("cannot write record: %s", err.Error())
		}
	}

	return nil
}

func (element *DeliveElement) toArray() []string {
	return []string{
		element.ref.NameJp,
		element.ref.NameEn,
		element.NameJp,
		element.NameEn,
		element.Description,
	}
}

// derive-elements-viewの書き込み
func (savedata *SaveData) WriteViewSegments(path string) error {
	// INFO: Writerの取得
	writer, cleanup, err := store.NewCsvWriter(path)
	if err != nil {
		return err
	}
	defer cleanup()

	// INFO: 書き込み
	defer writer.Flush() //内部バッファのフラッシュは必須
	writer.Write([]string{
		"項目名(JP)",
		"項目名(EN)",
		"区分値",
		"区分値名",
		"説明",
	})
	for _, elements := range savedata.Segments {
		if err := writer.Write(elements.toArray()); err != nil {
			return fmt.Errorf("cannot write record: %s", err.Error())
		}
	}

	return nil
}

func (element *Segment) toArray() []string {
	return []string{
		element.ref.NameJp,
		element.ref.NameEn,
		element.Value,
		element.Name,
		element.Description,
	}
}
