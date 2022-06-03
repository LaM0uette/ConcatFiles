package pkg

import (
	"ConcatFiles/loger"
	"io/ioutil"
	"os"
)

func GetCurrentDir() string {
	pwd, err := os.Getwd()
	if err != nil {
		loger.Error("Error with current dir:", err)
		os.Exit(1)
	}
	return pwd
}

func CreateNewFolder(path string) {
	if err := os.MkdirAll(path, os.ModePerm); err != nil {
		loger.Error("Erreur durant la création du dossier", err)
	}
}

func CopyFile(srcFile, dstFile string) {
	input, err := ioutil.ReadFile(srcFile)
	if err != nil {
		loger.Error("Erreur durant la copie du fichier (copier)", err)
	}

	err = ioutil.WriteFile(dstFile, input, 0644)
	if err != nil {
		loger.Error("Erreur durant la copie du fichier (coller)", err)
	}
}
