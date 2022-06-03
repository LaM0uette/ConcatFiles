package config

import (
	"log"
	"os/user"
	"path/filepath"
)

const (
	Name    = "FilesDIR"
	Author  = "LaM0uette"
	Version = "1.2.4"
)

var (
	DstPath     = filepath.Join(GetTempDir(), Name+"_Temp")
	LogsPath    = filepath.Join(DstPath, "logs")
	DumpsPath   = filepath.Join(DstPath, "dumps")
	ExportsPath = filepath.Join(DstPath, "exports")
)

func GetTempDir() string {
	temp, err := user.Current()
	if err != nil {
		log.Fatal(err)
	}

	return filepath.Join(temp.HomeDir)
}
