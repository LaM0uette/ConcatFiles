package pkg

import (
	"ConcatFile/loger"
	"FilesDIR/pkg"
	"fmt"
	"github.com/tealeg/xlsx"
	"path"
	"time"
)

var (
	Wb *xlsx.File
)

func CreateExcelFile() {

	Wb = xlsx.NewFile()

	err := Wb.Save(path.Join(pkg.GetCurrentDir(), fmt.Sprintf("Export_%v.xlsx", time.Now().Format("20060102150405"))))
	if err != nil {
		loger.Crash("Erreur lors de la création du fichier Excel", err)
	}
}
