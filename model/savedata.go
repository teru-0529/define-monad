package model

import (
	"errors"
	"fmt"
	"os"
	"slices"
	"strconv"

	"github.com/teru-0529/define-monad/store"
	"gopkg.in/yaml.v3"
)

type Dom string

var (
	UUID     Dom = "UUID"
	ID       Dom = "NOKEY"
	SEQUENCE Dom = "ID"
	ENUM     Dom = "区分値"
	CODE     Dom = "コード値"
	BOOL     Dom = "可否/フラグ"
	DATETIME Dom = "日時"
	DATE     Dom = "日付"
	TIME     Dom = "時間"
	INTEGER  Dom = "整数"
	NUMBER   Dom = "実数"
	STRING   Dom = "文字列"
	TEXT     Dom = "テキスト"
)

type SaveData struct {
	DataType       string          `yaml:"data_type"`
	Version        string          `yaml:"version"`
	Elements       []Element       `yaml:"elements"`
	DeliveElements []DeliveElement `yaml:"delive_elements"`
	Segments       []Segment       `yaml:"segments"`
}

type Element struct {
	NameJp      string  `yaml:"name_jp"`
	NameEn      string  `yaml:"name_en"`
	Domain      Dom     `yaml:"domain"`
	RegEx       *string `yaml:"reg_ex"`
	MinDigits   *int    `yaml:"min_digits"`
	MaxDigits   *int    `yaml:"max_digits"`
	MinValue    *int    `yaml:"min_value"`
	MaxValue    *int    `yaml:"max_value"`
	Example     string  `yaml:"example"`
	Description string  `yaml:"description"`
}

type DeliveElement struct {
	Origin      string   `yaml:"origin"`
	NameJp      string   `yaml:"name_jp"`
	NameEn      string   `yaml:"name_en"`
	Description string   `yaml:"description"`
	ref         *Element // 参照元項目
}

type Segment struct {
	Key         string   `yaml:"key"`
	Value       string   `yaml:"value"`
	Name        string   `yaml:"name"`
	Description string   `yaml:"description"`
	ref         *Element // 参照元項目
}

func NewSaveData(path string) (*SaveData, error) {
	// INFO: read
	file, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("cannot read file: %s", err.Error())
	}

	// INFO: unmarchal
	var savedata SaveData
	err = yaml.Unmarshal([]byte(file), &savedata)
	if err != nil {
		return nil, err
	}

	// INFO: originの設定
	for i := range savedata.DeliveElements {
		element := &savedata.DeliveElements[i]
		original, err := savedata.getElement(element.Origin)
		if err != nil {
			return nil, err
		}
		element.ref = original
	}
	for i := range savedata.Segments {
		element := &savedata.Segments[i]
		original, err := savedata.getElement(element.Key)
		if err != nil {
			return nil, err
		}
		element.ref = original
	}
	return &savedata, nil
}

// 項目オブジェクトの取得
func (savedata *SaveData) getElement(nameJp string) (*Element, error) {
	for i := range savedata.Elements {
		if savedata.Elements[i].NameJp == nameJp {
			return &savedata.Elements[i], nil
		}
	}
	return nil, errors.New("element not found")
}

// yamlファイルの書き込み
func (savedata *SaveData) Write(path string) error {
	// INFO: Encoderの取得
	encoder, cleanup, err := store.NewYamlEncorder(path)
	if err != nil {
		return err
	}
	defer cleanup()
	err = encoder.Encode(&savedata)
	if err != nil {
		return err
	}

	return nil
}

func (element Element) MarshalYAML() (interface{}, error) {
	if element.Domain == BOOL {
		example, _ := strconv.ParseBool(element.Example)
		return struct {
			NameJp      string  `yaml:"name_jp"`
			NameEn      string  `yaml:"name_en"`
			Domain      Dom     `yaml:"domain"`
			RegEx       *string `yaml:"reg_ex"`
			MinDigits   *int    `yaml:"min_digits"`
			MaxDigits   *int    `yaml:"max_digits"`
			MinValue    *int    `yaml:"min_value"`
			Maxvalue    *int    `yaml:"max_value"`
			Example     bool    `yaml:"example"`
			Description string  `yaml:"description"`
		}{
			NameJp:      element.NameJp,
			NameEn:      element.NameEn,
			Domain:      element.Domain,
			RegEx:       element.RegEx,
			MinDigits:   element.MinDigits,
			MaxDigits:   element.MaxDigits,
			MinValue:    element.MinValue,
			Maxvalue:    element.MaxValue,
			Example:     example,
			Description: element.Description,
		}, nil
	} else if slices.Contains([]Dom{INTEGER, SEQUENCE}, element.Domain) {
		example, _ := strconv.ParseInt(element.Example, 10, 64)
		return struct {
			NameJp      string  `yaml:"name_jp"`
			NameEn      string  `yaml:"name_en"`
			Domain      Dom     `yaml:"domain"`
			RegEx       *string `yaml:"reg_ex"`
			MinDigits   *int    `yaml:"min_digits"`
			MaxDigits   *int    `yaml:"max_digits"`
			MinValue    *int    `yaml:"min_value"`
			Maxvalue    *int    `yaml:"max_value"`
			Example     int64   `yaml:"example"`
			Description string  `yaml:"description"`
		}{
			NameJp:      element.NameJp,
			NameEn:      element.NameEn,
			Domain:      element.Domain,
			RegEx:       element.RegEx,
			MinDigits:   element.MinDigits,
			MaxDigits:   element.MaxDigits,
			MinValue:    element.MinValue,
			Maxvalue:    element.MaxValue,
			Example:     example,
			Description: element.Description,
		}, nil
	} else if slices.Contains([]Dom{NUMBER}, element.Domain) {
		example, _ := strconv.ParseFloat(element.Example, 64)
		return struct {
			NameJp      string  `yaml:"name_jp"`
			NameEn      string  `yaml:"name_en"`
			Domain      Dom     `yaml:"domain"`
			RegEx       *string `yaml:"reg_ex"`
			MinDigits   *int    `yaml:"min_digits"`
			MaxDigits   *int    `yaml:"max_digits"`
			MinValue    *int    `yaml:"min_value"`
			Maxvalue    *int    `yaml:"max_value"`
			Example     float64 `yaml:"example"`
			Description string  `yaml:"description"`
		}{
			NameJp:      element.NameJp,
			NameEn:      element.NameEn,
			Domain:      element.Domain,
			RegEx:       element.RegEx,
			MinDigits:   element.MinDigits,
			MaxDigits:   element.MaxDigits,
			MinValue:    element.MinValue,
			Maxvalue:    element.MaxValue,
			Example:     example,
			Description: element.Description,
		}, nil
	} else {
		return struct {
			NameJp      string  `yaml:"name_jp"`
			NameEn      string  `yaml:"name_en"`
			Domain      Dom     `yaml:"domain"`
			RegEx       *string `yaml:"reg_ex"`
			MinDigits   *int    `yaml:"min_digits"`
			MaxDigits   *int    `yaml:"max_digits"`
			MinValue    *int    `yaml:"min_value"`
			Maxvalue    *int    `yaml:"max_value"`
			Example     string  `yaml:"example"`
			Description string  `yaml:"description"`
		}{
			NameJp:      element.NameJp,
			NameEn:      element.NameEn,
			Domain:      element.Domain,
			RegEx:       element.RegEx,
			MinDigits:   element.MinDigits,
			MaxDigits:   element.MaxDigits,
			MinValue:    element.MinValue,
			Maxvalue:    element.MaxValue,
			Example:     element.Example,
			Description: element.Description,
		}, nil
	}
}
