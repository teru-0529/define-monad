/*
 */
package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

var inputFile string

// loadCmd represents the load command
var loadCmd = &cobra.Command{
	Use:   "load",
	Short: "load from yaml data and set excel sheet.",
	Long:  "load from yaml data and set excel sheet.",
	Run: func(cmd *cobra.Command, args []string) {
		fmt.Println(inputFile)
		fmt.Println("load called")
	},
}

func init() {
	// INFO:フラグ値を変数にBind
	loadCmd.Flags().StringVarP(&inputFile, "in", "I", "./save_data.yaml", "load save data path")
}
