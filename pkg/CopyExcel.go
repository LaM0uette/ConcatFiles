package pkg

import (
	"ConcatFiles/loger"
	"github.com/qax-os/excelize"
)

var (
	Wba *excelize.File
)

func (d *Data) CopyExcelFile(file string) {

	CopyFile(file, d.XlFile)
	DrawParam("GENERATION DE LA FICHE D'EXPORT:", "OK")

	Wba, err := excelize.OpenFile(d.XlFile)
	if err != nil {
		loger.Error("Erreur lors de l'ouverture de la fiche excel", err)
	}

	err = Wba.Save()
	if err != nil {
		loger.Error("Erreur lors de la sauvergarde du fichier Excel", err)
	}
}
