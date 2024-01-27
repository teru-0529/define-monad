/*
Copyright © 2024 Teruaki Sato <andrea.pirlo.0529@gmail.com>
*/
package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

// saveCmd represents the save command
var saveCmd = &cobra.Command{
	Use:   "save",
	Short: "Get savedata from excel sheet and write yaml.",
	Long:  "Get savedata from excel sheet and write yaml.",
	Run: func(cmd *cobra.Command, args []string) {
		fmt.Println(savedataPath)
		fmt.Println("save called")
	},
}

func init() {
}
