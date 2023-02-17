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
	case "jointureGraceLight":
		txtMode = "Grace Light"
	case "jointureEbp":
		txtMode = "Ebp"
	}

	pkg.DrawStart(txtMode)
	pkg.DrawSep("BUILD")

	//srcFile := "T:\\RIP FTTH\\GEOMAP\\5_EXPORTS\\GeoTools\\RIP24\\NRO38_PM8_BIMI\\DOE_70%\\1-V1\\DLG"
	srcFile := pkg.GetCurrentDir()
	dstFile := path.Join(srcFile, "__Concat__")

	timestamp := time.Now().Format("20060102150405")
	xlFile := path.Join(dstFile, fmt.Sprintf("__Export%s_%v.xlsm", txtMode, timestamp))

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
		pkg.CopyFile(path.Join(config.PathDocs, "DLG.qgs"), path.Join(d.DstFile, fmt.Sprintf("__DLG_%v.qgs", timestamp)))
		err := pkg.CopyDir(path.Join(config.PathDocs, "CARTO"), path.Join(d.DstFile, "CARTO"))
		if err != nil {
			return
		}
		d.ConcatCSVGrace()
	case "jointureGraceLight":
		d.CopyExcelFile(path.Join(config.PathXlsm, "MJGLight.xlsm"))
		pkg.CopyFile(path.Join(config.PathDocs, "DLG.qgs"), path.Join(d.DstFile, fmt.Sprintf("__DLG_%v.qgs", timestamp)))
		err := pkg.CopyDir(path.Join(config.PathDocs, "CARTO"), path.Join(d.DstFile, "CARTO"))
		if err != nil {
			return
		}
		d.ConcatCSVGraceLight()
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
