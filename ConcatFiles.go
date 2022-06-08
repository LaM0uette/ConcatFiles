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

	txtMode := ""
	switch *FlgMode {
	case "jointureGrace":
		txtMode = "Jointure Grace"
	}

	pkg.DrawStart(txtMode)
	pkg.DrawSep("BUILD")

	//pt := "T:\\RIP FTTH\\RIP FTTH 47\\2_Dossiers\\3_FTTH\\NRO_19\\5_ZAPM\\2_Plans\\3_DOE\\NRO_19_PM_06\\REC-DPR-47-NERA-NER6-01-V3"
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
