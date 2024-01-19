package model

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strconv"

	"github.com/samber/lo"
	"github.com/teru-0529/define-monad/store"
	"gopkg.in/yaml.v3"
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
	Domain      string  `yaml:"domain"`
	RegEx       *string `yaml:"reg_ex"`
	MinDigits   *int    `yaml:"min_digits"`
	MaxDigits   *int    `yaml:"max_digits"`
	MinValue    *int    `yaml:"min_value"`
	Maxvalue    *int    `yaml:"max_value"`
	Example     string  `yaml:"example"`
	Description string  `yaml:"description"`
}

// Example     interface{} `yaml:"example"`

type DeliveElement struct {
	Origin      string `yaml:"origin"`
	NameJp      string `yaml:"name_jp"`
	NameEn      string `yaml:"name_en"`
	Description string `yaml:"description"`
}

type Segment struct {
	Key         string `yaml:"key"`
	Value       string `yaml:"value"`
	Name        string `yaml:"name"`
	Description string `yaml:"description"`
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

	// fmt.Println(savedata) WARNING:
	return &savedata, nil
}

// yamlファイルの書き込み
func (savedata *SaveData) Write(path string) error {
	// INFO: encode
	var buffer bytes.Buffer
	yamlEncoder := yaml.NewEncoder(&buffer)
	yamlEncoder.SetIndent(2) // this is what you're looking for
	err := yamlEncoder.Encode(&savedata)
	if err != nil {
		return err
	}

	// INFO: create-directory
	dir := filepath.Dir(path)
	if _, err := os.Stat(dir); os.IsNotExist(err) {
		if err := os.MkdirAll(dir, 0777); err != nil {
			return fmt.Errorf("cannot create directory: %s", err.Error())
		}
	}

	// INFO: write-file
	_, err = buffer.WriteString(path)
	if err != nil {
		return fmt.Errorf("cannot write file: %s", err.Error())
	}

	return nil
}

// yamlファイルの書き込み
func (savedata *SaveData) WriteViewElements(path string) error {
	// INFO: Writerの取得
	writer, cleanup, err := store.NewWriter(path)
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
		"API(type)",
		"API(format)",
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

// func (element *Element) MarshalYAML() (interface{}, error) {
// 	fmt.Println(element.Example)

// 	return element.Example, nil
// }

func (element *Element) toArray() []string {
	return []string{
		element.NameJp,
		element.NameEn,
		element.Domain,
		element.Domain,
		element.Domain,
		lo.Ternary(lo.IsNil(element.RegEx), "", *element.RegEx),
		toIntStr(element.MinDigits),
		toIntStr(element.MaxDigits),
		toIntStr(element.MinValue),
		toIntStr(element.Maxvalue),
		element.Example,
		// element.toExampleStr(),
		element.Description,
	}
}

// func (element *Element) toExampleStr() string {
// 	res, err := gats.ToString(element.Example)
// 	if err != nil {
// 		return ""
// 	} else {
// 		return res
// 	}
// }

func toIntStr(val *int) string {
	return lo.TernaryF(
		lo.IsNil(val),
		func() string { return "" },
		func() string { return strconv.Itoa(*val) },
	)
}
