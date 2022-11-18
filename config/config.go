package config

import (
	"log"
	"os/user"
	"path/filepath"
)

const (
	Name    = "ConcatFiles"
	Author  = "LaM0uette"
	Version = "1.6.3"

	PathXlsm = "T:\\- 4 Suivi Appuis\\26_MACROS\\GeoMap\\ConcatFiles\\Docs"
	PathDocs = "T:\\- 11 Outils\\GeoMap\\Docs"
)

var (
	DstPath   = filepath.Join(GetTempDir(), Name+"_Temp")
	LogsPath  = filepath.Join(DstPath, "logs")
	DumpsPath = filepath.Join(DstPath, "dumps")
)

func GetTempDir() string {
	temp, err := user.Current()
	if err != nil {
		log.Fatal(err)
	}

	return filepath.Join(temp.HomeDir)
}
