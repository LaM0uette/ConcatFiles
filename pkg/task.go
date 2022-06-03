package pkg

import (
	"ConcatFile/loger"
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
		loger.Crash("Erreur durant la création du dossier", err)
	}
}
