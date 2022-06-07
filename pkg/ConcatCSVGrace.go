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
	"time"
)

type Fibre struct {
	FoCode    string
	FoNumTube string
	FoColor   string
}

type Cable struct {
	CbCode string
	CbEti  string
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

var (
	TFibre []Fibre
)

func (d *Data) ConcatCSVGrace() {

	DrawSep("CONCAT CSV GRACE")

	d.CopyCSV()
	DrawParam("COPIE DES CSV:", "OK")

	d.CountPositions()
	DrawParam("NOMBRE DE POSTIONS:", strconv.Itoa(d.NbrPos))

	DrawSep("LANCEMENT DE LA COMPILATION")
	d.AppendDatasInStructs()
	DrawParam("AJOUT DES DONNÉES DANS LES STRUCTS:", "OK")

	d.RunConcat(path.Join(d.DstFile, "t_position.csv"))

	err := Wb.Save(path.Join(d.DstFile, fmt.Sprintf("__Export_%v.xlsx", time.Now().Format("20060102150405"))))
	if err != nil {
		loger.Error("Erreur lors de la sauvergarde du fichier Excel", err)
	}
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

func (d *Data) CountPositions() {

	tPositionPath := path.Join(d.DstFile, "t_position.csv")
	CsvData := ReadCSV(tPositionPath)
	d.NbrPos = len(CsvData)
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

func (d *Data) AppendDatasInStructs() {
	AppendFibre(path.Join(d.DstFile, "t_fibre.csv"))
}

func AppendFibre(file string) {
	CsvFibre := ReadCSV(file)

	for _, val := range CsvFibre {
		fItem := Fibre{
			FoCode:    val[0],
			FoNumTube: val[4],
			FoColor:   val[8],
		}
		TFibre = append(TFibre, fItem)
	}
}

func (d *Data) RunConcat(file string) {
	CsvData := ReadCSV(file)

	Sht := Wb.Sheet["Export"]
	NbrTot := 0

	for r, val := range CsvData {

		PsCode, _ := Sht.Cell(r, 0)
		PsNum, _ := Sht.Cell(r, 1)
		Ps1, _ := Sht.Cell(r, 2)
		Ps2, _ := Sht.Cell(r, 6)
		PsCsCode, _ := Sht.Cell(r, 10)
		PsTiCode, _ := Sht.Cell(r, 13)
		PsType, _ := Sht.Cell(r, 15)
		PsFunc, _ := Sht.Cell(r, 16)
		PsState, _ := Sht.Cell(r, 17)
		PsPreaff, _ := Sht.Cell(r, 18)
		PsComment, _ := Sht.Cell(r, 19)
		PsCreaDate, _ := Sht.Cell(r, 20)
		PsMajDate, _ := Sht.Cell(r, 21)
		PsMajSrc, _ := Sht.Cell(r, 22)
		PsAbdDate, _ := Sht.Cell(r, 23)
		PsAbdSrc, _ := Sht.Cell(r, 24)

		PsCode.Value = val[0]
		PsNum.Value = val[1]
		Ps1.Value = val[2]
		Ps2.Value = val[3]
		PsCsCode.Value = val[4]
		PsTiCode.Value = val[5]
		PsType.Value = val[6]
		PsFunc.Value = val[7]
		PsState.Value = val[8]
		PsPreaff.Value = val[9]
		PsComment.Value = val[10]
		PsCreaDate.Value = val[11]
		PsMajDate.Value = val[12]
		PsMajSrc.Value = val[13]
		PsAbdDate.Value = val[14]
		PsAbdSrc.Value = val[15]

		NbrTot++
		loger.Void(fmt.Sprintf("%v/%v", NbrTot, d.NbrPos))
	}

	loger.Ok(fmt.Sprintf("%v positions concaténées", NbrTot))
}
