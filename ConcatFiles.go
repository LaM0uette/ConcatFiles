//go:generate goversioninfo -icon=ConcatFiles.ico -manifest=ConcatFiles.exe.manifest
package main

import (
	"ConcatFiles/config"
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
	//jointureEbp
	flag.Parse()

	txtMode := ""
	switch *FlgMode {
	case "jointureGrace":
		txtMode = "Grace"
	case "jointureEbp":
		txtMode = "Ebp"
	}

	pkg.DrawStart(txtMode)
	pkg.DrawSep("BUILD")

	srcFile := "C:\\Users\\XD5965\\OneDrive - EQUANS\\Bureau\\REC-DLG-47-CAST-ANT1-01-V2"
	//srcFile := pkg.GetCurrentDir()
	dstFile := path.Join(srcFile, "__Concat__")
	xlFile := path.Join(dstFile, fmt.Sprintf("__Export%s_%v.xlsm", txtMode, time.Now().Format("20060102150405")))

	d := pkg.Data{
		SrcFile: srcFile,
		DstFile: dstFile,
		XlFile:  xlFile,
	}

	pkg.DrawParam("CREATION DU DOSSIER:", "OK")
	pkg.CreateNewFolder(d.DstFile)

	switch *FlgMode {
	case "jointureGrace":
		d.CopyExcelFile(path.Join(config.PathXlsm, "MJG.xlsm"))
		d.ConcatCSVGrace()
	case "jointureEbp":
		d.CopyExcelFile(path.Join(config.PathXlsm, "MJEbp.xlsm"))
		d.ConcatCSVEbp()
	}

	pkg.DrawSep(" FIN ")
	pkg.DrawEnd()

	rgb.GreenB.Print("Appuyer sur Entrée pour quitter...")
	_, err := bufio.NewReader(os.Stdin).ReadBytes('\n')
	if err != nil {
		loger.Crash("Crash :", err)
	}
}
