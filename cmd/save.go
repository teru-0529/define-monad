/*
Copyright © 2024 Teruaki Sato <andrea.pirlo.0529@gmail.com>
*/
package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
	"github.com/teru-0529/define-monad/model"
)

// saveCmd represents the save command
var saveCmd = &cobra.Command{
	Use:   "save",
	Short: "Get savedata from excel sheet and write yaml.",
	Long:  "Get savedata from excel sheet and write yaml.",
	RunE: func(cmd *cobra.Command, args []string) error {

		monad, err := model.FromExcel(excelPath)
		if err != nil {
			return err
		}

		monad.Version = version
		// fmt.Println(monad)
		monad.Write("./monad.yaml") // FIXME:最終的にはsavedataに書き出す

		fmt.Println("***command[save] completed.")
		return nil
	},
}

func init() {
}
