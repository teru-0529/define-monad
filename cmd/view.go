/*
Copyright © 2024 Teruaki Sato <andrea.pirlo.0529@gmail.com>
*/
package cmd

import (
	"fmt"
	"path/filepath"

	"github.com/spf13/cobra"
	"github.com/teru-0529/define-monad/model"
)

var outDir string
var elementsFile string
var deriveElementsFile string
var segmentsFile string

// viewCmd represents the view command
var viewCmd = &cobra.Command{
	Use:   "view",
	Short: "Generate tsv data for sphinx from savedata.",
	Long:  "Generate tsv data for sphinx from savedata.",
	RunE: func(cmd *cobra.Command, args []string) error {

		monad, err := model.NewSaveData(savedataPath)
		if err != nil {
			return err
		}

		elementsPath := filepath.Join(outDir, elementsFile)
		monad.WriteViewElements(elementsPath)
		deriveElementsPath := filepath.Join(outDir, deriveElementsFile)
		monad.WriteViewDeriveElements(deriveElementsPath)
		segmentsPath := filepath.Join(outDir, segmentsFile)
		monad.WriteViewSegments(segmentsPath)

		fmt.Printf("input yaml file: [%s]\n", filepath.ToSlash(filepath.Clean(savedataPath)))
		fmt.Printf("output tsv file(elements): [%s]\n", filepath.ToSlash(elementsPath))
		fmt.Printf("output tsv file(derive-elements): [%s]\n", filepath.ToSlash(deriveElementsPath))
		fmt.Printf("output tsv file(segments): [%s]\n", filepath.ToSlash(segmentsPath))
		fmt.Println("***command[view] completed.")
		return nil
	},
}

func init() {
	viewCmd.Flags().StringVarP(&outDir, "out-dir", "O", "./view", "output directory")
	viewCmd.Flags().StringVarP(&elementsFile, "elements-file", "", "view_elements.tsv", "output elements file name")
	viewCmd.Flags().StringVarP(&deriveElementsFile, "derive-elements-file", "", "view_derive_elements.tsv", "output derive-elements file name")
	viewCmd.Flags().StringVarP(&segmentsFile, "segments-file", "", "view_segments.tsv", "output segment file name")
}
