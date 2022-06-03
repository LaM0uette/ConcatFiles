package pkg

import (
	"ConcatFile/loger"
	"fmt"
	"github.com/tealeg/xlsx"
	"path"
	"time"
)

var (
	Wb *xlsx.File
)

func CreateExcelFile() {

	DrawParam("GENERATION DE LA FICHE D'EXPORT")

	Wb = xlsx.NewFile()

	_, err := Wb.AddSheet("Export")
	if err != nil {
		loger.Crash("Erreur lors de la création de l'onglet Export", err)
	}

	err = Wb.Save(path.Join(GetCurrentDir(), fmt.Sprintf("Export_%v.xlsx", time.Now().Format("20060102150405"))))
	if err != nil {
		loger.Crash("Erreur lors de la sauvergarde du fichier Excel", err)
	}
}
