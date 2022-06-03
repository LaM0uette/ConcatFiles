package main

import "ConcatFile/pkg"

func main() {

	pkg.DrawStart()

	pkg.DrawSep("BUILD")
	pkg.CreateExcelFile()

	pkg.DrawSep(" FIN ")
	pkg.DrawEnd()
}
