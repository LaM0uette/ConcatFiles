package pkg

import "path"

func ConcatCSVGrace() {

	DrawSep("CONCAT CSV GRACE")

	DrawParam("CREATION DU DOSSIER")
	CreateNewFolder(path.Join(GetCurrentDir(), "csv"))

	DrawParam("DEPLACEMENT DES CSV")
}
