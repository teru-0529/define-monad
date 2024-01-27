package model

import (
	"fmt"
	"strings"

	"github.com/samber/lo"
	"github.com/teru-0529/define-monad/store"
)

// types-ddlの書き込み
func (savedata *SaveData) WriteTypesDdl(path string) error {
	// INFO: Fileの取得
	file, cleanup, err := store.NewFile(path)
	if err != nil {
		return err
	}
	defer cleanup()

	// INFO: header書き込み
	file.WriteString("-- Enum Type DDL\n\n")

	// INFO: ENUMの項目のみ処理
	for _, element := range lo.Filter(savedata.Elements, func(item Element, index int) bool { return item.Domain == ENUM }) {
		nameJp, nameEn := element.NameJp, snakeCase(element.NameEn)

		file.WriteString(fmt.Sprintf("-- %s\nDROP TYPE IF EXISTS %s;\n", nameJp, nameEn))
		file.WriteString(fmt.Sprintf("CREATE TYPE %s AS enum (\n", nameEn))

		// INFO: 該当する項目を抽出
		items := lo.FilterMap(savedata.Segments, func(item Segment, index int) (string, bool) {
			if item.Key == nameJp {
				return fmt.Sprintf("  '%s'", item.Value), true
			} else {
				return "", false
			}
		})
		file.WriteString(strings.Join(items, ",\n"))

		file.WriteString("\n);\n\n")
	}

	return nil
}

// スネークケース変換(あえて独自に実装：数値を大文字と同じ(数値の前にアンダーバー)とする仕様)
func snakeCase(camel string) string {
	if camel == "" {
		return camel
	}

	delimiter := "_"
	sLen := len(camel)
	var snake string
	for i, current := range camel {
		if i > 0 && i+1 < sLen {
			if current >= '0' && current <= 'Z' {
				prev := camel[i-1]
				if prev >= 'a' && prev <= 'z' {
					snake += delimiter
				}
			}
		}
		snake += string(current)
	}

	snake = strings.ToLower(snake)
	return snake
}
