package loger

import (
	"FilesDIR/config"
	"FilesDIR/rgb"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"time"
)

var (
	// logs
	ui    *log.Logger
	ok    *log.Logger
	nok   *log.Logger
	errr  *log.Logger
	crash *log.Logger

	semicolon *log.Logger
)

const (
	preOk    = "[ -OK- ]"
	preNok   = "[ -NOK- ]"
	preErrr  = "[ -ERROR- ]"
	preCrash = "[ -CRASH- ]"
)

func init() {

	createTempFiles()

	logFile, err := os.OpenFile(filepath.Join(config.LogsPath, fmt.Sprintf("SLog_%v.log", time.Now().Format("20060102150405"))), os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0666)
	if err != nil {
		log.Fatal(err)
	}

	dumpFile, err := os.OpenFile(filepath.Join(config.DumpsPath, fmt.Sprintf("Dump_%v.csv", time.Now().Format("20060102150405"))), os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0666)
	if err != nil {
		log.Fatal(err)
	}

	ui = log.New(logFile, "", 0)
	ok = log.New(logFile, preOk+" ", log.Ltime|log.Lmsgprefix)
	nok = log.New(logFile, preNok+" ", log.Ltime|log.Lmsgprefix)
	errr = log.New(logFile, preErrr+" ", log.Ltime|log.Lmsgprefix)
	crash = log.New(logFile, preCrash+" ", log.Ltime|log.Lmsgprefix)

	semicolon = log.New(dumpFile, "", 0)
}

//...
// Log
func Ui(v ...any) {
	ui.Print(v...)
}

func Void(msg string) {
	fmt.Print("\r", rgb.GreenB.Sprint(msg))
}

func Ok(msg string) {
	ok.Print(msg)
	fmt.Print("\r", rgb.GreenBg.Sprint(preOk), rgb.GreenB.Sprint(" ", msg), "\n")
}

func Nok(msg string) {
	nok.Print(msg)
	fmt.Print("\r", rgb.RedBg.Sprint(preNok), rgb.RedB.Sprint(" ", msg), "\n")
}

func Error(msg string, err any) {
	errr.Print(msg, " ", err)
	fmt.Print(rgb.RedBg.Sprint(preErrr), rgb.RedB.Sprint(" ", msg), rgb.RedB.Sprint(" ", err), "\n")
}

func Errorinf(msg string, err any) {
	errr.Print(msg, " ", err)
	fmt.Print("\r", rgb.RedBg.Sprint(preErrr), rgb.RedB.Sprint(" ", msg), rgb.RedB.Sprint(" ", err), "\n")
}

func Crash(msg string, err any) {
	crash.Print(msg, " ", err)
	fmt.Print(rgb.RedBg.Sprint(preCrash), rgb.RedBg.Sprint(" ", msg), rgb.RedB.Sprint(" ", err), "\n")
	os.Exit(1)
}

//...
// Dump
func Semicolon(v ...any) {
	semicolon.Println(v...)
}

// ...
// Init func
func createTempFiles() {
	folders := []string{
		"logs", "dumps", "exports", "blacklist", "whitelist",
	}

	var paths []string

	for _, f := range folders {
		paths = append(paths, filepath.Join(config.DstPath, f))
	}

	for _, p := range paths {
		err := os.MkdirAll(p, os.ModePerm)
		if err != nil {
			log.Fatal(err)
		}
	}

	createFile(filepath.Join(config.DstPath, "words.txt"))
	createFile(filepath.Join(filepath.Join(config.DstPath, "blacklist"), "__ALL__.txt"))
	createFile(filepath.Join(filepath.Join(config.DstPath, "whitelist"), "__ALL__.txt"))
}

func createFile(file string) {
	var _, err = os.Stat(file)

	if os.IsNotExist(err) {
		var file, err = os.Create(file)
		if err != nil {
			log.Fatal(err)
		}
		defer func(file *os.File) {
			err := file.Close()
			if err != nil {
				log.Fatal(err)
			}
		}(file)
	}
}
