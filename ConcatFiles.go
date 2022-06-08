//go:generate goversioninfo -icon=ConcatFiles.ico -manifest=ConcatFiles.exe.manifest
package main

import (
	"ConcatFiles/loger"
	"ConcatFiles/pkg"
	"ConcatFiles/rgb"
	"bufio"
	"flag"
	"os"
	"path"
)

func main() {

	FlgMode := flag.String("m", "jointureGrace", "Mode de compilation")
	flag.Parse()

	pkg.DrawStart()
	pkg.DrawSep("BUILD")

	//pt := "C:\\Users\\XD5965\\OneDrive - EQUANS\\Bureau\\DLG"
	pt := pkg.GetCurrentDir()

	d := pkg.Data{
		SrcFile: pt,
		DstFile: path.Join(pt, "__Concat__"),
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
