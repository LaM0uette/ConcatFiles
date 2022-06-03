package main

import (
	"ConcatFile/pkg"
	"flag"
)

func main() {

	FlgMode := flag.String("m", "jointureGrace", "Mode de compilation")
	flag.Parse()

	pkg.DrawStart()

	pkg.DrawSep("BUILD")
	pkg.CreateExcelFile()

	switch *FlgMode {
	case "jointureGrace":

	}

	pkg.DrawSep(" FIN ")
	pkg.DrawEnd()
}
