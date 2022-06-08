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

	pt := "C:\\Users\\XD5965\\OneDrive - EQUANS\\Bureau\\REC-DPR-47-NERA-NER6-01-V3"
	//pt := pkg.GetCurrentDir()

	d := pkg.Data{
		SrcFile: pt,
		DstFile: path.Join(pt, "__Concat__"),
		XlFile:  path.Join(pt, "__Concat__", fmt.Sprintf("__Export_%v.xlsm", time.Now().Format("20060102150405"))),
	}

	pkg.DrawParam("CREATION DU DOSSIER:", "OK")
	pkg.CreateNewFolder(d.DstFile)

	d.CreateExcelFile()

	switch *FlgMode {
	case "jointureGrace":
		d.ConcatCSVGrace()
	}

	pkg.DrawSep(" FIN ")
	pkg.DrawEnd()

	rgb.GreenB.Print("Appuyer sur Entrée pour quitter...")
	_, err := bufio.NewReader(os.Stdin).ReadBytes('\n')
	if err != nil {
		loger.Crash("Crash :", err)
	}
}
