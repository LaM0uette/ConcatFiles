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
	FoCbCode  string
}

type Cable struct {
	CbCode string
	CbEti  string
}

type Cassette struct {
	CsCode   string
	CsNum    string
	CsBpCode string
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
	NameTCable    = "t_cable.csv"
	NameTCassette = "t_cassette.csv"
	NameTEbp      = "t_ebp.csv"
	NameTFibre    = "t_fibre.csv"
	NameTPosition = "t_position.csv"
	NameTTiroir   = "t_tiroir.csv"

	TFibre    []Fibre
	TCable    []Cable
	TCassette []Cassette
	TEbp      []Ebp
	TTirroir  []Tirroir
)

func (d *Data) ConcatCSVGrace() {

	DrawSep("CONCAT CSV GRACE")

	d.CopyCSV()
	DrawParam("COPIE DES CSV:", "OK")

	d.CountPositions()
	DrawParam("NOMBRE DE POSTIONS:", strconv.Itoa(d.NbrPos))

	d.AppendDatasInStructs()
	DrawParam("AJOUT DES DONNÉES DANS LES STRUCTS:", "OK")

	DrawSep("LANCEMENT DE LA COMPILATION")
	d.RunConcat(path.Join(d.DstFile, NameTPosition))
	setHeaderWb()

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

	CopyFile(path.Join(dlgPath, NameTCable), path.Join(d.DstFile, NameTCable))
	CopyFile(path.Join(dlgPath, NameTCassette), path.Join(d.DstFile, NameTCassette))
	CopyFile(path.Join(dlgPath, NameTEbp), path.Join(d.DstFile, NameTEbp))
	CopyFile(path.Join(dlgPath, NameTFibre), path.Join(d.DstFile, NameTFibre))
	CopyFile(path.Join(dlgPath, NameTPosition), path.Join(d.DstFile, NameTPosition))
	CopyFile(path.Join(dlgPath, NameTTiroir), path.Join(d.DstFile, NameTTiroir))
}

func (d *Data) CountPositions() {

	tPositionPath := path.Join(d.DstFile, NameTPosition)
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
	AppendFibre(path.Join(d.DstFile, NameTFibre))
	AppendCable(path.Join(d.DstFile, NameTCable))
	AppendCassette(path.Join(d.DstFile, NameTCassette))
	AppendEbp(path.Join(d.DstFile, NameTEbp))
	AppendTirroir(path.Join(d.DstFile, NameTTiroir))
}
func AppendFibre(file string) {
	Csv := ReadCSV(file)

	for _, val := range Csv {
		Item := Fibre{
			FoCode:    val[0],
			FoNumTube: val[4],
			FoColor:   val[8],
			FoCbCode:  val[2],
		}
		TFibre = append(TFibre, Item)
	}
}
func AppendCable(file string) {
	Csv := ReadCSV(file)

	for _, val := range Csv {
		Item := Cable{
			CbCode: val[0],
			CbEti:  val[2],
		}
		TCable = append(TCable, Item)
	}
}
func AppendCassette(file string) {
	Csv := ReadCSV(file)

	for _, val := range Csv {
		Item := Cassette{
			CsCode:   val[0],
			CsNum:    val[3],
			CsBpCode: val[2],
		}
		TCassette = append(TCassette, Item)
	}
}
func AppendEbp(file string) {
	Csv := ReadCSV(file)

	for _, val := range Csv {
		Item := Ebp{
			BpCode: val[0],
			BpEti:  val[1],
		}
		TEbp = append(TEbp, Item)
	}
}
func AppendTirroir(file string) {
	Csv := ReadCSV(file)

	for _, val := range Csv {
		Item := Tirroir{
			TiCode: val[0],
			TiEti:  val[2],
		}
		TTirroir = append(TTirroir, Item)
	}
}

func setHeaderWb() {
	header := []string{"ps_code", "ps_numero", "ps_1", "fo_numtub", "fo_color", "cb_etiquet", "ps_2", "fo_numtub", "fo_color", "cb_etiquet", "ps_cs_code", "cs_num", "bp_etiquet", "ps_ti_code", "ti_etiquet", "ps_type", "ps_fonct", "ps_etat", "ps_preaff", "ps_comment", "ps_creadat", "ps_majdate", "ps_majsrc", "ps_abddate", "ps_abdsrc"}

	Sht := Wb.Sheet["Export"]
	for i, v := range header {
		cell, _ := Sht.Cell(0, i)
		cell.Value = v
	}
}

func (d *Data) RunConcat(file string) {
	CsvData := ReadCSV(file)

	Sht := Wb.Sheet["Export"]
	NbrTot := 0

	for r, val := range CsvData {
		fo1 := getDataFibre(val[2])   //ps1 et cb1
		fo2 := getDataFibre(val[3])   //ps2 et cb2
		cs := getDataCassette(val[4]) //psCsCode
		bp := getDataEbp(cs[1])       //csCode

		PsCode, _ := Sht.Cell(r, 0)
		PsNum, _ := Sht.Cell(r, 1)
		Ps1, _ := Sht.Cell(r, 2)
		FoNumTube1, _ := Sht.Cell(r, 3)
		FoColor1, _ := Sht.Cell(r, 4)
		CbEti1, _ := Sht.Cell(r, 5)
		Ps2, _ := Sht.Cell(r, 6)
		FoNumTube2, _ := Sht.Cell(r, 7)
		FoColor2, _ := Sht.Cell(r, 8)
		CbEti2, _ := Sht.Cell(r, 9)
		PsCsCode, _ := Sht.Cell(r, 10)
		CsNum, _ := Sht.Cell(r, 11)
		BpEti, _ := Sht.Cell(r, 12)
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
		FoNumTube1.Value = fo1[0]
		FoColor1.Value = fo1[1]
		CbEti1.Value = fo1[2]
		Ps2.Value = val[3]
		FoNumTube2.Value = fo2[0]
		FoColor2.Value = fo2[1]
		CbEti2.Value = fo2[2]
		PsCsCode.Value = val[4]
		CsNum.Value = cs[0]
		BpEti.Value = bp
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

func getDataFibre(ps string) []string {
	for _, data := range TFibre {
		if ps == data.FoCode {
			return []string{data.FoNumTube, data.FoColor, data.FoCbCode}
		}
	}
	return []string{"", "", ""}
}

func getDataCassette(cs string) []string {
	for _, data := range TCassette {
		if cs == data.CsCode {
			return []string{data.CsNum, data.CsBpCode}
		}
	}
	return []string{"", ""}
}

func getDataEbp(bp string) string {
	for _, data := range TEbp {
		if bp == data.BpCode {
			return data.BpEti
		}
	}
	return ""
}
