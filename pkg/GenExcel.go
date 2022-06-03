package pkg

import (
	"ConcatFile/loger"
	"FilesDIR/pkg"
	"github.com/tealeg/xlsx"
)

var (
	Wb *xlsx.File
)

func CreateExcelFile() {

	Wb = xlsx.NewFile()

	err := Wb.Save(pkg.GetCurrentDir())
	if err != nil {
		loger.Crash("Erreur lors de la création du fichier Excel", err)
	}
}
