package pkg

import (
	"ConcatFile/loger"
	"os"
	"path"
	"path/filepath"
	"strings"
)

func (d *Data) ConcatCSVGrace() {

	DrawSep("CONCAT CSV GRACE")

	DrawParam("CREATION DU DOSSIER")
	CreateNewFolder(path.Join(GetCurrentDir(), "csv"))

	d.GetFolderDLG()

	DrawParam("COPIE DES CSV")
	dlgPath := path.Join(d.SrcFile, d.GetFolderDLG())
	CopyFile(path.Join(dlgPath, "t_cable.csv"), d.DstFile)
	CopyFile(path.Join(dlgPath, "t_cassette.csv"), d.DstFile)
	CopyFile(path.Join(dlgPath, "t_ebp.csv"), d.DstFile)
	CopyFile(path.Join(dlgPath, "t_fibre.csv"), d.DstFile)
	CopyFile(path.Join(dlgPath, "t_position.csv"), d.DstFile)
	CopyFile(path.Join(dlgPath, "t_tiroir.csv"), d.DstFile)
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
		loger.Crash("Crash pendant le listing des dossiers", err)
	}

	return dlg
}
