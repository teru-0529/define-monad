package elements

import (
	"bytes"
	"fmt"
	"slices"
	"strconv"
	"strings"

	"github.com/samber/lo"
	"github.com/teru-0529/define-monad/v3/store"
	"gopkg.in/yaml.v3"
)

type ApiElement struct {
	ApiType     string   `yaml:"type"`
	ApiFormat   string   `yaml:"format,omitempty"`
	RegEx       *string  `yaml:"pattern,omitempty"`
	Enum        []string `yaml:"enum,omitempty"`
	MinDigits   *int     `yaml:"minLength,omitempty"`
	MaxDigits   *int     `yaml:"maxLength,omitempty"`
	MinValue    *int     `yaml:"minimum,omitempty"`
	MaxValue    *int     `yaml:"maximum,omitempty"`
	Description string   `yaml:"description"`
	Example     string   `yaml:"example"`
}

// WARNING:削除
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

// api-elementsの書き込み
func (savedata *SaveData) WriteApiElements(path string) error {
	// INFO: Fileの取得
	file, cleanup, err := store.NewFile(path)
	if err != nil {
		return err
	}
	defer cleanup()

	// INFO: Elementの項目を順次処理
	for _, element := range savedata.Elements {

		file.WriteString(fmt.Sprintf("# %s\n%s:\n", element.NameJp, element.NameEn))

		apiElement := ApiElement{
			ApiType:     element.apiType(),
			ApiFormat:   element.apiFormat(),
			RegEx:       element.RegEx,
			Enum:        savedata.segmentValues(element.NameJp),
			MinDigits:   element.MinDigits,
			MaxDigits:   element.MaxDigits,
			MinValue:    element.MinValue,
			MaxValue:    element.MaxValue,
			Description: element.Description,
			Example:     element.Example,
		}

		file.WriteString(apiElement.toYamlStr())
	}

	return nil
}

// WARNING:削除
// db-elementsの書き込み
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

// 区分値を配列で返す
func (savedata *SaveData) segmentValues(key string) []string {
	return lo.FilterMap(savedata.Segments, func(item Segment, _ int) (string, bool) {
		if item.Key == key {
			return item.Value, true
		} else {
			return "", false
		}
	})
}

// yaml文字列の出力
func (apiElement *ApiElement) toYamlStr() string {
	// INFO: encode（インデントは2）
	var b bytes.Buffer
	encoder := yaml.NewEncoder(&b)
	encoder.SetIndent(2)
	encoder.Encode(apiElement)

	// INFO: 要素なのでインデントを追加（一度配列化してMap処理）
	// 文字列にしてから改行コードで分割し配列化
	arr := strings.Split(b.String(), "\n")
	// 最終要素を削除（最終行に改行コードがあるため）
	arr = append([]string{}, arr[:len(arr)-1]...)
	// インデント追加
	arr = lo.Map(arr, func(item string, index int) string { return fmt.Sprintf("  %s", item) })
	// 改行コードで連結しなおし、最後に改行コードを追加
	return fmt.Sprintf("%s\n\n", strings.Join(arr, "\n"))
}

// toYaml
func (element ApiElement) MarshalYAML() (interface{}, error) {
	if element.ApiType == "boolean" {
		example, _ := strconv.ParseBool(element.Example)
		return struct {
			ApiType     string   `yaml:"type"`
			ApiFormat   string   `yaml:"format,omitempty"`
			RegEx       *string  `yaml:"pattern,omitempty"`
			Enum        []string `yaml:"enum,omitempty"`
			MinDigits   *int     `yaml:"minLength,omitempty"`
			MaxDigits   *int     `yaml:"maxLength,omitempty"`
			MinValue    *int     `yaml:"minimum,omitempty"`
			MaxValue    *int     `yaml:"maximum,omitempty"`
			Description string   `yaml:"description"`
			Example     bool     `yaml:"example"`
		}{
			ApiType:     element.ApiType,
			ApiFormat:   element.ApiFormat,
			RegEx:       element.RegEx,
			Enum:        element.Enum,
			MinDigits:   element.MinDigits,
			MaxDigits:   element.MaxDigits,
			MinValue:    element.MinValue,
			MaxValue:    element.MaxValue,
			Description: element.Description,
			Example:     example,
		}, nil
	} else if element.ApiType == "integer" {
		example, _ := strconv.ParseInt(element.Example, 10, 64)
		return struct {
			ApiType     string   `yaml:"type"`
			ApiFormat   string   `yaml:"format,omitempty"`
			RegEx       *string  `yaml:"pattern,omitempty"`
			Enum        []string `yaml:"enum,omitempty"`
			MinDigits   *int     `yaml:"minLength,omitempty"`
			MaxDigits   *int     `yaml:"maxLength,omitempty"`
			MinValue    *int     `yaml:"minimum,omitempty"`
			MaxValue    *int     `yaml:"maximum,omitempty"`
			Description string   `yaml:"description"`
			Example     int64    `yaml:"example"`
		}{
			ApiType:     element.ApiType,
			ApiFormat:   element.ApiFormat,
			RegEx:       element.RegEx,
			Enum:        element.Enum,
			MinDigits:   element.MinDigits,
			MaxDigits:   element.MaxDigits,
			MinValue:    element.MinValue,
			MaxValue:    element.MaxValue,
			Description: element.Description,
			Example:     example,
		}, nil
	} else if element.ApiType == "number" {
		example, _ := strconv.ParseFloat(element.Example, 64)
		return struct {
			ApiType     string   `yaml:"type"`
			ApiFormat   string   `yaml:"format,omitempty"`
			RegEx       *string  `yaml:"pattern,omitempty"`
			Enum        []string `yaml:"enum,omitempty"`
			MinDigits   *int     `yaml:"minLength,omitempty"`
			MaxDigits   *int     `yaml:"maxLength,omitempty"`
			MinValue    *int     `yaml:"minimum,omitempty"`
			MaxValue    *int     `yaml:"maximum,omitempty"`
			Description string   `yaml:"description"`
			Example     float64  `yaml:"example"`
		}{
			ApiType:     element.ApiType,
			ApiFormat:   element.ApiFormat,
			RegEx:       element.RegEx,
			Enum:        element.Enum,
			MinDigits:   element.MinDigits,
			MaxDigits:   element.MaxDigits,
			MinValue:    element.MinValue,
			MaxValue:    element.MaxValue,
			Description: element.Description,
			Example:     example,
		}, nil
	} else {
		return struct {
			ApiType     string   `yaml:"type"`
			ApiFormat   string   `yaml:"format,omitempty"`
			RegEx       *string  `yaml:"pattern,omitempty"`
			Enum        []string `yaml:"enum,omitempty"`
			MinDigits   *int     `yaml:"minLength,omitempty"`
			MaxDigits   *int     `yaml:"maxLength,omitempty"`
			MinValue    *int     `yaml:"minimum,omitempty"`
			MaxValue    *int     `yaml:"maximum,omitempty"`
			Description string   `yaml:"description"`
			Example     string   `yaml:"example"`
		}{
			ApiType:     element.ApiType,
			ApiFormat:   element.ApiFormat,
			RegEx:       element.RegEx,
			Enum:        element.Enum,
			MinDigits:   element.MinDigits,
			MaxDigits:   element.MaxDigits,
			MinValue:    element.MinValue,
			MaxValue:    element.MaxValue,
			Description: element.Description,
			Example:     element.Example,
		}, nil
	}
}

// apiスキーマタイプ
func (element *Element) apiType() string {
	if slices.Contains([]Dom{SEQUENCE, INTEGER}, element.Domain) {
		return "integer"
	} else if element.Domain == BOOL {
		return "boolean"
	} else if element.Domain == NUMBER {
		return "number"
	} else {
		return "string"
	}
}

// apiスキーマフォーマット
func (element *Element) apiFormat() string {
	if element.Domain == UUID {
		return "uuid"
	} else if element.Domain == DATE {
		return "date"
	} else if element.Domain == DATETIME {
		return "date-time"
	} else {
		return ""
	}
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
	} else if element.Domain == UUID {
		return "uuid"
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

// WARNING:削除
// db制約(リダイレクトメソッド)
func (element *Element) constraint() *string {
	return element._constraint(element.nameEnSnake())
}

// WARNING:削除
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
	} else if element.MinValue != nil && element.MaxValue != nil {
		// 最小値・最大値の両方が設定されている
		result := fmt.Sprintf("(%d <= %s AND %s <= %d)", *element.MinValue, nameEnSnake, nameEnSnake, *element.MaxValue)
		return &result
	} else if element.MinValue != nil {
		// 最小値が設定されている
		result := fmt.Sprintf("(%s >= %d)", nameEnSnake, *element.MinValue)
		return &result
	} else if element.MaxValue != nil {
		// 最大値が設定されている
		result := fmt.Sprintf("(%s <= %d)", nameEnSnake, *element.MaxValue)
		return &result
	}
	return nil
}

// WARNING:削除
// DDLでNotNullを強制するかどうか
func (element *Element) mustNotNull() bool {
	return slices.Contains([]Dom{ENUM, BOOL}, element.Domain)
}

// WARNING:削除
// DDLに記載するデフォルトが文字列扱いかどうか
func (element *Element) isDefaultStr() bool {
	return slices.Contains([]Dom{ID, ENUM, CODE, STRING, TEXT}, element.Domain)
}

// スネークケースの名称
func (element *DeliveElement) nameEnSnake() string {
	return snakeCase(element.NameEn)
}

// WARNING:削除
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
