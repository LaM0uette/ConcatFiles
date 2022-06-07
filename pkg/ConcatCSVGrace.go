package pkg

import (
	"ConcatFiles/loger"
	"bufio"
	"encoding/csv"
	"fmt"
	"os"
	"path"
	"path/filepath"
	"strconv"
	"strings"
)

type Position struct {
	PsCode     string
	PsNum      int
	Ps1        string
	Ps2        string
	PsCsCode   string
	PsTiCode   string
	PsFunc     string
	PsState    string
	PsPreaff   string
	PsComment  string
	PsCreaDate string
	PsMajDate  string
	PsMajSrc   string
	PsAbdDate  string
	PsAbdSrc   string
}

type Fibre struct {
	FoCode     string
	FoNumTube1 int
	FoColor1   float32
	FoNumTube2 int
	FoColor2   float32
}

type Cable struct {
	CbCode string
	CbEti1 string
	CbEti2 string
}

type Cassette struct {
	CsCode string
	CbNum  string
}

type Ebp struct {
	BpCode string
	BpEti  string
}

type Tirroir struct {
	TiCode string
	TiEti  string
}

func (d *Data) ConcatCSVGrace() {

	DrawSep("CONCAT CSV GRACE")

	d.CopyCSV()
	DrawParam("COPIE DES CSV:", "OK")

	DrawParam("NOMBRE DE POSTIONS:", strconv.Itoa(d.CountPositions()))

	DrawSep("LANCEMENT DE LA COMPILATION")
	d.AppendStructData()
}

func (d *Data) GetFolderDLG() string {
	var dlg string

	err := filepath.Walk(d.SrcFile, func(path string, fileInfo os.FileInfo, err error) error {
		if fileInfo.IsDir() && strings.Contains(fileInfo.Name(), "-DLG-") {
			dlg = fileInfo.Name()
			return nil
		}
		return nil
	})

	if err != nil {
		loger.Error("Error pendant le listing des dossiers", err)
	}

	return dlg
}

func (d *Data) CopyCSV() {
	dlgPath := path.Join(d.SrcFile, d.GetFolderDLG())

	CopyFile(path.Join(dlgPath, "t_cable.csv"), path.Join(d.DstFile, "t_cable.csv"))
	CopyFile(path.Join(dlgPath, "t_cassette.csv"), path.Join(d.DstFile, "t_cassette.csv"))
	CopyFile(path.Join(dlgPath, "t_ebp.csv"), path.Join(d.DstFile, "t_ebp.csv"))
	CopyFile(path.Join(dlgPath, "t_fibre.csv"), path.Join(d.DstFile, "t_fibre.csv"))
	CopyFile(path.Join(dlgPath, "t_position.csv"), path.Join(d.DstFile, "t_position.csv"))
	CopyFile(path.Join(dlgPath, "t_tiroir.csv"), path.Join(d.DstFile, "t_tiroir.csv"))
}

func (d *Data) CountPositions() int {

	tPositionPath := path.Join(d.DstFile, "t_position.csv")
	CsvData := ReadCSV(tPositionPath)
	return len(CsvData)
}

func (d *Data) AppendStructData() {

	TPosition := path.Join(d.DstFile, "t_position.csv")

	p := &Position{}
	p.FillingPosition(TPosition)
}

func ReadCSV(file string) [][]string {

	CsvFile, err := os.Open(file)
	if err != nil {
		loger.Error(fmt.Sprintf("Error lors de l'ouverture de %s:", file), err)
	}
	defer func(tPosition *os.File) {
		err := tPosition.Close()
		if err != nil {
			loger.Error(fmt.Sprintf("Error lors de la fermeture de %s:", file), err)
		}
	}(CsvFile)

	reader := csv.NewReader(bufio.NewReader(CsvFile))
	reader.Comma = ';'
	reader.LazyQuotes = true

	CsvData, err := reader.ReadAll()
	if err != nil {
		loger.Error(fmt.Sprintf("Error lors de la lecture des données de %s:", file), err)
	}
	return CsvData
}

func (p *Position) FillingPosition(file string) {

	CsvData := ReadCSV(file)
	fmt.Println(CsvData)
}
