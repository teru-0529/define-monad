/*
Copyright © 2024 Teruaki Sato <andrea.pirlo.0529@gmail.com>
*/
package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

var outputFile string

// saveCmd represents the save command
var saveCmd = &cobra.Command{
	Use:   "save",
	Short: "set data from excel sheet and write yaml data.",
	Long:  "set data from excel sheet and write yaml data.",
	Run: func(cmd *cobra.Command, args []string) {
		fmt.Println(outputFile)
		fmt.Println("save called")
	},
}

func init() {
	// INFO:フラグ値を変数にBind
	saveCmd.Flags().StringVarP(&outputFile, "out", "O", "./save_data.yaml", "save data path")
}
