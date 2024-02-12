package elements

import (
	"fmt"
	"slices"
	"strings"

	"github.com/samber/lo"
)

// toExcel(sht-elements)
func (savedata *SaveData) ToExcelElements() []string {
	return lo.Map(savedata.Elements, func(item Element, index int) string { return item.toExcel() })
}

func (element *Element) toExcel() string {
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
	return strings.Join(ary, "\t")
}

// toExcel(sht-derive-elements)
func (savedata *SaveData) ToExcelDeriveElements() []string {
	return lo.Map(savedata.DeliveElements, func(item DeliveElement, index int) string { return item.toExcel() })
}

func (element *DeliveElement) toExcel() string {
	ary := []string{
		element.ref.NameJp,
		element.NameJp,
		element.NameEn,
		element.Description,
	}
	return strings.Join(ary, "\t")
}

// toExcel(sht-derive-elements)
func (savedata *SaveData) ToExcelSegments() []string {
	return lo.Map(savedata.Segments, func(item Segment, index int) string { return item.toExcel() })
}

func (element *Segment) toExcel() string {
	ary := []string{
		element.ref.NameJp,
		element.Value,
		element.Name,
		element.Description,
	}
	return strings.Join(ary, "\t")
}

// toTableElements
func (savedata *SaveData) TotableElements() []string {
	result := []string{}
	for _, element := range savedata.Elements {
		result = append(result, element.toTable())
	}
	for _, element := range savedata.DeliveElements {
		result = append(result, element.toTable())
	}
	return result
}

// WARNING:メソッド名
func (element *Element) toTable() string {
	ary := []string{
		element.NameJp,
		element.nameEnSnake(),
		element.dbModel(),
		element.constraint_(),
		element.mustNotNull_(),
		element.isDefaultStr_(),
		element.Description,
		"0",
		"",
	}
	return strings.Join(ary, "\t")
}

// WARNING:メソッド名
func (element *DeliveElement) toTable() string {
	ary := []string{
		element.NameJp,
		element.nameEnSnake(),
		element.ref.dbModel(),
		element.constraint_(),
		element.ref.mustNotNull_(),
		element.ref.isDefaultStr_(),
		element.Description,
		"1",
		element.ref.NameJp,
	}
	return strings.Join(ary, "\t")
}

// WARNING:メソッド名
// db制約(リダイレクトメソッド)
func (element *Element) constraint_() string {
	return element._constraint_(element.nameEnSnake())
}

// WARNING:メソッド名
// db制約
func (element *Element) _constraint_(nameEnSnake string) string {
	if element.RegEx != nil {
		// 正規表現が設定されている
		return fmt.Sprintf("(%s ~* '%s')", nameEnSnake, *element.RegEx)
	} else if element.MinDigits != nil {
		// 最小桁数が設定されている
		return fmt.Sprintf("(LENGTH(%s) >= %d)", nameEnSnake, *element.MinDigits)
	} else if element.MinValue != nil && element.MaxValue != nil {
		// 最小値・最大値の両方が設定されている
		return fmt.Sprintf("(%d <= %s AND %s <= %d)", *element.MinValue, nameEnSnake, nameEnSnake, *element.MaxValue)
	} else if element.MinValue != nil {
		// 最小値が設定されている
		return fmt.Sprintf("(%s >= %d)", nameEnSnake, *element.MinValue)
	} else if element.MaxValue != nil {
		// 最大値が設定されている
		return fmt.Sprintf("(%s <= %d)", nameEnSnake, *element.MaxValue)
	}
	return ""
}

// WARNING:メソッド名
// DDLでNotNullを強制するかどうか
func (element *Element) mustNotNull_() string {
	if slices.Contains([]Dom{ENUM, BOOL}, element.Domain) {
		return "true"
	} else {
		return "false"
	}
}

// WARNING:メソッド名
// DDLに記載するデフォルトが文字列扱いかどうか
func (element *Element) isDefaultStr_() string {
	if slices.Contains([]Dom{ID, ENUM, CODE, STRING, TEXT}, element.Domain) {
		return "true"
	} else {
		return "false"
	}
}

// WARNING:メソッド名
// db制約(リダイレクトメソッド)
func (element *DeliveElement) constraint_() string {
	return element.ref._constraint_(element.nameEnSnake())
}
