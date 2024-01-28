package model

import (
	"fmt"
	"slices"
	"strings"

	"github.com/samber/lo"
	"github.com/teru-0529/define-monad/store"
)

type TableElement struct {
	NameJp          string  `yaml:"name_jp"`
	NameEn          string  `yaml:"name_en"`
	DbModel         string  `yaml:"db_model"`
	Constraint      *string `yaml:"constraint"`
	MustNotNull     bool    `yaml:"must_not_null"`
	IsStringDefault bool    `yaml:"is_string_default"`
	Description     string  `yaml:"description"`
	IsOrigin        bool    `yaml:"is_origin"`
	Origin          *string `yaml:"origin"`
	Dummy           *string `yaml:"dummy-row"` //FIXME:DB側が修正されるまでのダミー
}

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

// types-ddlの書き込み
func (savedata *SaveData) WriteTableElements(path string) error {
	tableElements := []TableElement{}

	for _, element := range savedata.Elements {
		tableElements = append(tableElements, TableElement{
			NameJp:          element.NameJp,
			NameEn:          element.nameEnSnake(),
			DbModel:         element.dbModel(),
			Constraint:      element.constraint(),
			MustNotNull:     element.mustNotNull(),
			IsStringDefault: element.isDefaultStr(),
			Description:     element.Description,
			IsOrigin:        true,
			Origin:          nil,
		})
	}

	for _, element := range savedata.DeliveElements {
		tableElements = append(tableElements, TableElement{
			NameJp:          element.NameJp,
			NameEn:          element.nameEnSnake(),
			DbModel:         element.ref.dbModel(),
			Constraint:      element.constraint(),
			MustNotNull:     element.ref.mustNotNull(),
			IsStringDefault: element.ref.isDefaultStr(),
			Description:     element.Description,
			IsOrigin:        false,
			Origin:          &element.ref.NameJp,
		})
	}

	// INFO: Encoderの取得
	encoder, cleanup, err := store.NewYamlEncorder(path)
	if err != nil {
		return err
	}
	defer cleanup()

	// INFO: 書き込み
	err = encoder.Encode(&tableElements)
	if err != nil {
		return err
	}
	return nil
}

// スネークケースの名称
func (element *Element) nameEnSnake() string {
	return snakeCase(element.NameEn)
}

// db制約
func (element *Element) dbModel() string {
	if element.Domain == SEQUENCE {
		return "serial"
	} else if element.Domain == BOOL {
		return "boolean"
	} else if element.Domain == INTEGER {
		return "integer"
	} else if element.Domain == NUMBER {
		return "numeric"
	} else if element.Domain == TEXT {
		return "text"
	} else if element.Domain == DATE {
		return "date"
	} else if element.Domain == DATETIME {
		return "timestamp"
	} else if element.Domain == TIME {
		return "varchar(5)"
	} else if element.Domain == ENUM {
		return element.nameEnSnake()
	} else {
		return fmt.Sprintf("varchar(%d)", *element.MaxDigits)
	}
}

// db制約(リダイレクトメソッド)
func (element *Element) constraint() *string {
	return element._constraint(element.nameEnSnake())
}

// db制約
func (element *Element) _constraint(nameEnSnake string) *string {
	if element.RegEx != nil {
		// 正規表現が設定されている
		result := fmt.Sprintf("(%s ~* '%s')", nameEnSnake, *element.RegEx)
		return &result
	} else if element.MinDigits != nil {
		// 最小桁数が設定されている
		result := fmt.Sprintf("(LENGTH(%s) >= %d)", nameEnSnake, *element.MinDigits)
		return &result
	} else if element.MinValue != nil && element.Maxvalue != nil {
		// 最小値・最大値の両方が設定されている
		result := fmt.Sprintf("(%d <= %s AND %s <= %d)", *element.MinValue, nameEnSnake, nameEnSnake, *element.Maxvalue)
		return &result
	} else if element.MinValue != nil {
		// 最小値が設定されている
		result := fmt.Sprintf("(%s >= %d)", nameEnSnake, *element.MinValue)
		return &result
	} else if element.Maxvalue != nil {
		// 最大値が設定されている
		result := fmt.Sprintf("(%s <= %d)", nameEnSnake, *element.Maxvalue)
		return &result
	}
	return nil
}

// DDLでNotNullを強制するかどうか
func (element *Element) mustNotNull() bool {
	return slices.Contains([]Dom{ENUM, BOOL}, element.Domain)
}

// DDLに記載するデフォルトが文字列扱いかどうか
func (element *Element) isDefaultStr() bool {
	return slices.Contains([]Dom{ID, ENUM, CODE, STRING, TEXT}, element.Domain)
}

// スネークケースの名称
func (element *DeliveElement) nameEnSnake() string {
	return snakeCase(element.NameEn)
}

// db制約(リダイレクトメソッド)
func (element *DeliveElement) constraint() *string {
	return element.ref._constraint(element.nameEnSnake())
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
