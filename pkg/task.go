package pkg

import (
	"ConcatFiles/loger"
	"bufio"
	"encoding/csv"
	"fmt"
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

func ReadCSV(file string) [][]string {

	CsvFile, err := os.Open(file)
	if err != nil {
		loger.Error(fmt.Sprintf("Error lors de l'ouverture de %s:", file), err)
	}
	defer func(tPosition *os.File) {
		err := tPosition.Close()
		if err != nil {
			loger.Error(fmt.Sprintf("Error lors de la fermeture de %s:", file), err)
		}
	}(CsvFile)

	reader := csv.NewReader(bufio.NewReader(CsvFile))
	reader.Comma = ';'
	reader.LazyQuotes = true

	CsvData, err := reader.ReadAll()
	if err != nil {
		loger.Error(fmt.Sprintf("Error lors de la lecture des données de %s:", file), err)
	}
	return CsvData
}
