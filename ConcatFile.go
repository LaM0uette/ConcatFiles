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

	d := pkg.Data{
		SrcFile: "T:\\RIP FTTH\\GEOMAP\\5_EXPORTS\\RIP40\\NRO7_PM3_NISY\\DOE_1_100%\\V12\\DLG",
		DstFile: path.Join(pkg.GetCurrentDir(), "__Concat__"),
	}

	d.CreateExcelFile()

	switch *FlgMode {
	case "jointureGrace":
		d.ConcatCSVGrace()
	}

	pkg.DrawSep(" FIN ")
	pkg.DrawEnd()
}

//SrcFile: pkg.GetCurrentDir(),
