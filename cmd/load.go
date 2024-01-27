/*
 */
package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

// loadCmd represents the load command
var loadCmd = &cobra.Command{
	Use:   "load",
	Short: "Read savedata from yaml and set excel sheet.",
	Long:  "Read savedata from yaml and set excel sheet.",
	Run: func(cmd *cobra.Command, args []string) {
		fmt.Println(savedataPath)
		fmt.Println("***command[load] completed.")
	},
}

func init() {
}
