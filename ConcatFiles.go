//go:generate goversioninfo -icon=ConcatFiles.ico -manifest=ConcatFiles.exe.manifest
package main

import (
	"ConcatFiles/loger"
	"ConcatFiles/pkg"
	"ConcatFiles/rgb"
	"bufio"
	"flag"
	"fmt"
	"os"
	"path"
	"time"
)

func main() {
	FlgMode := flag.String("m", "jointureGrace", "Mode de compilation")
	flag.Parse()

	txtMode := ""
	switch *FlgMode {
	case "jointureGrace":
		txtMode = "Jointure Grace"
	}

	pkg.DrawStart(txtMode)
	pkg.DrawSep("BUILD")

	srcFile := "C:\\Users\\XD5965\\OneDrive - EQUANS\\Bureau\\REC-DPR-47-NERA-NER6-01-V3"
	//srcFile := pkg.GetCurrentDir()
	dstFile := path.Join(srcFile, "__Concat__")
	xlFile := path.Join(dstFile, fmt.Sprintf("__Export_%v.xlsm", time.Now().Format("20060102150405")))

	d := pkg.Data{
		SrcFile: srcFile,
		DstFile: dstFile,
		XlFile:  xlFile,
	}

	pkg.DrawParam("CREATION DU DOSSIER:", "OK")
	pkg.CreateNewFolder(d.DstFile)

	switch *FlgMode {
	case "jointureGrace":
		d.CopyExcelFile("T:\\- 4 Suivi Appuis\\26_MACROS\\GO\\ConcatFiles\\Docs\\MJG.xlsm")
		d.ConcatCSVGrace()
	}

	//d.CreateExcelFile()

	pkg.DrawSep(" FIN ")
	pkg.DrawEnd()

	rgb.GreenB.Print("Appuyer sur Entrée pour quitter...")
	_, err := bufio.NewReader(os.Stdin).ReadBytes('\n')
	if err != nil {
		loger.Crash("Crash :", err)
	}
}
