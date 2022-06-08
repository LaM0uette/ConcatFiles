package pkg

import (
	"ConcatFiles/loger"
	"github.com/tealeg/xlsx"
)

type Data struct {
	SrcFile string
	DstFile string
	XlFile  string
	NbrPos  int
}

var (
	Wb *xlsx.File
)

func (d *Data) CreateExcelFile() {

	DrawParam("GENERATION DE LA FICHE D'EXPORT:", "OK")

	CopyFile("T:\\- 4 Suivi Appuis\\26_MACROS\\GO\\ConcatFiles\\Docs\\MacroJointureGrace.xlsm", d.XlFile)

	Wb, err := xlsx.OpenFile(d.XlFile)
	if err != nil {
		loger.Error("Erreur lors de l'ouverture de la fiche MacroJointureGrace", err)
	}

	_, err = Wb.AddSheet("Export")
	if err != nil {
		loger.Error("Erreur lors de la création de l'onglet Export", err)
	}

	//err = Wb.Save(path.Join(d.DstFile, fmt.Sprintf("__Export_%v.xlsx", time.Now().Format("20060102150405"))))
	//if err != nil {
	//	loger.Error("Erreur lors de la sauvergarde du fichier Excel", err)
	//}
}
