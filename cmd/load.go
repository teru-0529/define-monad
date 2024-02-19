/*
 */
package cmd

import (
	"fmt"
	"slices"

	"github.com/spf13/cobra"
	"github.com/teru-0529/define-monad/v3/model/elements"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
)

type ShtType string

var (
	SHT_ELEMENTS        ShtType = "elements"
	SHT_DERIVE_ELEMENTS ShtType = "derive-elements"
	SHT_SEGMENTS        ShtType = "segments"
)

var shtTypeStr string

// loadCmd represents the load command
var loadCmd = &cobra.Command{
	Use:   "load",
	Short: "Read savedata from yaml and set excel sheet.",
	Long:  "Read savedata from yaml and set excel sheet.",
	RunE: func(cmd *cobra.Command, args []string) error {

		// INFO: 出力タイプが想定された種類ではない場合エラーを返す
		var shtType ShtType = ShtType(shtTypeStr)
		if !slices.Contains([]ShtType{SHT_ELEMENTS, SHT_DERIVE_ELEMENTS, SHT_SEGMENTS}, shtType) {
			return fmt.Errorf("input `sht-type` is not found: [%s]", shtTypeStr)
		}

		// INFO: save-dataの読込み
		monad, err := elements.New(savedataPath)
		if err != nil {
			// エラーにせず空のデータを返す
			monad = &elements.SaveData{}
		}

		enc := japanese.ShiftJIS.NewEncoder()

		if shtType == SHT_ELEMENTS {
			for _, rec := range monad.ToExcelElements() {
				// sjisに変換
				sjis, _, _ := transform.String(enc, rec)
				fmt.Println(sjis)
			}

		} else if shtType == SHT_DERIVE_ELEMENTS {
			for _, rec := range monad.ToExcelDeriveElements() {
				// sjisに変換
				sjis, _, _ := transform.String(enc, rec)
				fmt.Println(sjis)
			}

		} else if shtType == SHT_SEGMENTS {
			for _, rec := range monad.ToExcelSegments() {
				// sjisに変換
				sjis, _, _ := transform.String(enc, rec)
				fmt.Println(sjis)
			}

		}

		return nil
	},
}

func init() {
	loadCmd.Flags().StringVarP(&shtTypeStr, "sheet-type", "S", "", "setting sheet type(one of 'elements','derive-elements','segments')")

	loadCmd.MarkFlagRequired("sheet-type")

}
