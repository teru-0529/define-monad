package elements

import (
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
		element.Ref.NameJp,
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
		element.Ref.NameJp,
		element.Value,
		element.Name,
		element.Description,
	}
	return strings.Join(ary, "\t")
}
