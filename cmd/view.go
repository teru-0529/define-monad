/*
Copyright © 2024 Teruaki Sato <andrea.pirlo.0529@gmail.com>
*/
package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
	"github.com/teru-0529/define-monad/model"
)

// viewCmd represents the view command
var viewCmd = &cobra.Command{
	Use:   "view",
	Short: "Generate tsv data for sphinx from savedata.",
	Long:  "Generate tsv data for sphinx from savedata.",
	RunE: func(cmd *cobra.Command, args []string) error {

		monad, err := model.NewSaveData("./testdata/save_data.yaml")
		if err != nil {
			return err
		}

		monad.Write("./monad.yaml")
		monad.WriteViewElements("./view_elements.tsv")
		fmt.Println("view called")
		return nil
	},
}

func init() {

	// Here you will define your flags and configuration settings.

	// Cobra supports Persistent Flags which will work for this command
	// and all subcommands, e.g.:
	// viewCmd.PersistentFlags().String("foo", "", "A help for foo")

	// Cobra supports local flags which will only run when this command
	// is called directly, e.g.:
	// viewCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
}
