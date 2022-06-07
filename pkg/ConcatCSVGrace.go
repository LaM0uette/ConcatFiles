package pkg

import (
	"ConcatFiles/loger"
	"bufio"
	"encoding/csv"
	"os"
	"path"
	"path/filepath"
	"strconv"
	"strings"
)

func (d *Data) ConcatCSVGrace() {

	DrawSep("CONCAT CSV GRACE")

	d.CopyCSV()
	DrawParam("COPIE DES CSV:", "OK")

	DrawParam("NOMBRE DE POSTIONS:", strconv.Itoa(d.CountPositions()))

	DrawSep("LANCEMENT DE LA COMPILATION")
}

func (d *Data) GetFolderDLG() string {
	var dlg string

	err := filepath.Walk(d.SrcFile, func(path string, fileInfo os.FileInfo, err error) error {
		if fileInfo.IsDir() && strings.Contains(fileInfo.Name(), "-DLG-") {
			dlg = fileInfo.Name()
			return nil
		}
		return nil
	})

	if err != nil {
		loger.Error("Error pendant le listing des dossiers", err)
	}

	return dlg
}

func (d *Data) CopyCSV() {
	dlgPath := path.Join(d.SrcFile, d.GetFolderDLG())

	CopyFile(path.Join(dlgPath, "t_cable.csv"), path.Join(d.DstFile, "t_cable.csv"))
	CopyFile(path.Join(dlgPath, "t_cassette.csv"), path.Join(d.DstFile, "t_cassette.csv"))
	CopyFile(path.Join(dlgPath, "t_ebp.csv"), path.Join(d.DstFile, "t_ebp.csv"))
	CopyFile(path.Join(dlgPath, "t_fibre.csv"), path.Join(d.DstFile, "t_fibre.csv"))
	CopyFile(path.Join(dlgPath, "t_position.csv"), path.Join(d.DstFile, "t_position.csv"))
	CopyFile(path.Join(dlgPath, "t_tiroir.csv"), path.Join(d.DstFile, "t_tiroir.csv"))
}

func (d *Data) CountPositions() int {

	tPositionPath := path.Join(d.DstFile, "t_position.csv")

	tPosition, err := os.Open(tPositionPath)
	if err != nil {
		loger.Error("Error lors de l'ouverture de la t_position:", err)
	}
	defer func(tPosition *os.File) {
		err := tPosition.Close()
		if err != nil {
			loger.Error("Error lors de la fermeture de la t_position:", err)
		}
	}(tPosition)

	reader := csv.NewReader(bufio.NewReader(tPosition))
	reader.Comma = ';'
	reader.LazyQuotes = true

	csvLines, err := reader.ReadAll()
	if err != nil {
		return 0
	}

	return len(csvLines)
}
