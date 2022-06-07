package pkg

import (
	"ConcatFiles/loger"
	"os"
	"path"
	"path/filepath"
	"strings"
)

func (d *Data) ConcatCSVGrace() {

	DrawSep("CONCAT CSV GRACE")

	DrawParam("COPIE DES CSV")
	dlgPath := path.Join(d.SrcFile, d.GetFolderDLG())
	CopyFile(path.Join(dlgPath, "t_cable.csv"), path.Join(d.DstFile, "t_cable.csv"))
	CopyFile(path.Join(dlgPath, "t_cassette.csv"), path.Join(d.DstFile, "t_cassette.csv"))
	CopyFile(path.Join(dlgPath, "t_ebp.csv"), path.Join(d.DstFile, "t_ebp.csv"))
	CopyFile(path.Join(dlgPath, "t_fibre.csv"), path.Join(d.DstFile, "t_fibre.csv"))
	CopyFile(path.Join(dlgPath, "t_position.csv"), path.Join(d.DstFile, "t_position.csv"))
	CopyFile(path.Join(dlgPath, "t_tiroir.csv"), path.Join(d.DstFile, "t_tiroir.csv"))
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
