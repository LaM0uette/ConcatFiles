//go:generate goversioninfo -icon=ConcatFile.ico -manifest=ConcatFile.exe.manifest
package main

import (
	"ConcatFile/pkg"
	"flag"
	"path"
)

func main() {

	FlgMode := flag.String("m", "jointureGrace", "Mode de compilation")
	flag.Parse()

	pkg.DrawStart()
	pkg.DrawSep("BUILD")

	pt := "T:\\RIP FTTH\\GEOMAP\\5_EXPORTS\\RIP40\\NRO7_PM3_NISY\\DOE_1_100%\\V12\\DLG"
	//pt := SrcFile: pkg.GetCurrentDir(),

	d := pkg.Data{
		SrcFile: pt,
		DstFile: path.Join(pt, "__Concat__"),
	}

	pkg.DrawParam("CREATION DU DOSSIER")
	pkg.CreateNewFolder(d.DstFile)

	d.CreateExcelFile()

	switch *FlgMode {
	case "jointureGrace":
		d.ConcatCSVGrace()
	}

	pkg.DrawSep(" FIN ")
	pkg.DrawEnd()
}
