/*
Copyright © 2024 Teruaki Sato <andrea.pirlo.0529@gmail.com>
*/
package cmd

import (
	"fmt"
	"path/filepath"
	"slices"

	"github.com/spf13/cobra"
	"github.com/teru-0529/define-monad/model"
)

type OutType string

var (
	TYPE_DDL    OutType = "type-ddl"
	API_ELEMENT OutType = "api-element"
	DB_ELEMENT  OutType = "db-element"
)

var distTypeStr string
var distFile string

// outputCmd represents the output command
var outputCmd = &cobra.Command{
	Use:   "output",
	Short: "output file in each format from savedata.",
	Long:  "output file in each format from savedata.",
	RunE: func(cmd *cobra.Command, args []string) error {

		// INFO: 出力タイプが想定された種類ではない場合エラーを返す
		var distType OutType = OutType(distTypeStr)
		if !slices.Contains([]OutType{TYPE_DDL, API_ELEMENT, DB_ELEMENT}, distType) {
			return fmt.Errorf("input `dist-type` is not found: [%s]", distTypeStr)
		}

		// INFO: save-dataの読込み
		monad, err := model.NewSaveData(savedataPath)
		if err != nil {
			return err
		}
		fmt.Printf("input yaml file: [%s]\n", filepath.ToSlash(filepath.Clean(savedataPath)))

		if distType == TYPE_DDL {
			monad.WriteTypesDdl(distFile)
			fmt.Println("output type: type-ddl")
			fmt.Printf("output sql file: [%s]\n", filepath.ToSlash(distFile))

		} else if distType == API_ELEMENT {
			monad.WriteApiElements(distFile)
			fmt.Println("output type: api-element")
			fmt.Printf("output yaml file: [%s]\n", filepath.ToSlash(distFile))

		} else if distType == DB_ELEMENT {
			err := monad.WriteTableElements(distFile)
			if err != nil {
				return err
			}
			fmt.Println("output type: db-element")
			fmt.Printf("output yaml file: [%s]\n", filepath.ToSlash(distFile))

		}

		fmt.Println("***command[output] completed.")
		return nil
	},
}

func init() {
	outputCmd.Flags().StringVarP(&distTypeStr, "dist-type", "T", "", "output file type(one of 'type-ddl','api-element','db-element')")
	outputCmd.Flags().StringVarP(&distFile, "dist-file", "F", "", "output file name")

	outputCmd.MarkFlagRequired("dist-type")
	outputCmd.MarkFlagRequired("dist-file")
}
