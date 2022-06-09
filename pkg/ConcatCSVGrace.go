package pkg

import (
	"ConcatFiles/loger"
	"fmt"
	"github.com/qax-os/excelize"
	"github.com/tealeg/xlsx"
	"os"
	"path"
	"path/filepath"
	"strconv"
	"strings"
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
type PositionErr struct {
	PsCode string
}

var (
	Wba        *excelize.File
	fitToWidth excelize.FitToWidth

	NameTCable    = "t_cable.csv"
	NameTCassette = "t_cassette.csv"
	NameTEbp      = "t_ebp.csv"
	NameTFibre    = "t_fibre.csv"
	NameTPosition = "t_position.csv"
	NameTTiroir   = "t_tiroir.csv"

	TFibre       []Fibre
	TCable       []Cable
	TCassette    []Cassette
	TEbp         []Ebp
	TTirroir     []Tirroir
	TPositionErr []PositionErr
)

func (d *Data) ConcatCSVGrace() {
	//Wba, err := excelize.OpenFile(d.XlFile)
	//if err != nil {
	//	loger.Error("Erreur lors de l'ouverture de la fiche excel", err)
	//}

	Wba, _ = excelize.OpenFile(d.XlFile)

	DrawSep("CONCAT CSV GRACE")

	d.copyCSV()
	DrawParam("COPIE DES CSV:", "OK")

	d.countPositions()
	DrawParam("NOMBRE DE POSTIONS:", strconv.Itoa(d.NbrPos))

	d.checkIfErrExist()

	d.appendDatasInStructs()
	DrawParam("AJOUT DES DONNÉES DANS LES STRUCTS:", "OK")

	DrawSep("COMPILATION")
	d.runConcat(path.Join(d.DstFile, NameTPosition))
	setFormatingWb()

	//err := Wba.AddVBAProject("T:\\- 4 Suivi Appuis\\26_MACROS\\GO\\ConcatFiles\\Docs\\MJG.bin")
	//if err != nil {
	//	loger.Error("Erreur pendant l'ajout du module VBA'", err)
	//}

	_ = Wba.SetCellValue("MACRO", "A3", d.DstFile)

	_ = Wba.SetAppProps(&excelize.AppProperties{
		Application:       "Microsoft Excel",
		ScaleCrop:         true,
		DocSecurity:       2,
		Company:           "LaM0uette",
		LinksUpToDate:     true,
		HyperlinksChanged: true,
		AppVersion:        "16.0000",
	})

	if err := Wba.SaveAs(d.XlFile); err != nil {
		loger.Error("Erreur pendant la sauvegarde du fichier Excel:", err)
	}
}

func (d *Data) copyCSV() {
	dlgPath := path.Join(d.SrcFile, d.getFolderDLG())

	CopyFile(path.Join(dlgPath, NameTCable), path.Join(d.DstFile, NameTCable))
	CopyFile(path.Join(dlgPath, NameTCassette), path.Join(d.DstFile, NameTCassette))
	CopyFile(path.Join(dlgPath, NameTEbp), path.Join(d.DstFile, NameTEbp))
	CopyFile(path.Join(dlgPath, NameTFibre), path.Join(d.DstFile, NameTFibre))
	CopyFile(path.Join(dlgPath, NameTPosition), path.Join(d.DstFile, NameTPosition))
	CopyFile(path.Join(dlgPath, NameTTiroir), path.Join(d.DstFile, NameTTiroir))
}

func (d *Data) countPositions() {
	tPositionPath := path.Join(d.DstFile, NameTPosition)
	CsvData := ReadCSV(tPositionPath)
	d.NbrPos = len(CsvData)
}

func (d *Data) checkIfErrExist() {
	f := path.Join(d.SrcFile, d.getDLGErr())

	if !FileExist(f) {
		return
	}

	WbErr, err := xlsx.OpenFile(f)
	if err != nil {
		loger.Error("Erreur à l'ouverture du fichier", err)
		return
	}

	for _, sheet := range WbErr.Sheets {
		if sheet.Name == "POSITION" {
			counter := 0

			for i := 0; i < sheet.MaxRow; i++ {
				cell, _ := sheet.Cell(i, 1)
				if strings.Contains(cell.Value, "PS") {
					p := PositionErr{PsCode: cell.Value}
					TPositionErr = append(TPositionErr, p)
					counter++
				}
			}

			DrawParam("NOMBRES DE POSITIONS EN ERREURS:", strconv.Itoa(counter))
			break
		}
	}

}

func (d *Data) appendDatasInStructs() {
	appendFibre(path.Join(d.DstFile, NameTFibre))
	appendCable(path.Join(d.DstFile, NameTCable))
	appendCassette(path.Join(d.DstFile, NameTCassette))
	appendEbp(path.Join(d.DstFile, NameTEbp))
	appendTirroir(path.Join(d.DstFile, NameTTiroir))
}
func appendFibre(file string) {
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
func appendCable(file string) {
	Csv := ReadCSV(file)

	for _, val := range Csv {
		Item := Cable{
			CbCode: val[0],
			CbEti:  val[2],
		}
		TCable = append(TCable, Item)
	}
}
func appendCassette(file string) {
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
func appendEbp(file string) {
	Csv := ReadCSV(file)

	for _, val := range Csv {
		Item := Ebp{
			BpCode: val[0],
			BpEti:  val[1],
		}
		TEbp = append(TEbp, Item)
	}
}
func appendTirroir(file string) {
	Csv := ReadCSV(file)

	for _, val := range Csv {
		Item := Tirroir{
			TiCode: val[0],
			TiEti:  val[2],
		}
		TTirroir = append(TTirroir, Item)
	}
}

func (d *Data) runConcat(file string) {
	CsvData := ReadCSV(file)

	NbrTot := 0

	for r, val := range CsvData {
		//fo1 := getDataFibre(val[2])   //ps1 et cb1
		//fo2 := getDataFibre(val[3])   //ps2 et cb2
		//cs := getDataCassette(val[4]) //psCsCode
		//bp := getDataEbp(cs[1])       //csCode
		//ti := getDataTirroir(val[5])  //tiCode

		//PsCode, _ := Sht.Cell(r, 0)
		//PsNum, _ := Sht.Cell(r, 1)
		//Ps1, _ := Sht.Cell(r, 2)
		//FoNumTube1, _ := Sht.Cell(r, 3)
		//FoColor1, _ := Sht.Cell(r, 4)
		//CbEti1, _ := Sht.Cell(r, 5)
		//Ps2, _ := Sht.Cell(r, 6)
		//FoNumTube2, _ := Sht.Cell(r, 7)
		//FoColor2, _ := Sht.Cell(r, 8)
		//CbEti2, _ := Sht.Cell(r, 9)
		//PsCsCode, _ := Sht.Cell(r, 10)
		//CsNum, _ := Sht.Cell(r, 11)
		//BpEti, _ := Sht.Cell(r, 12)
		//PsTiCode, _ := Sht.Cell(r, 13)
		//TiEti, _ := Sht.Cell(r, 14)
		//PsType, _ := Sht.Cell(r, 15)
		//PsFunc, _ := Sht.Cell(r, 16)
		//PsState, _ := Sht.Cell(r, 17)
		//PsPreaff, _ := Sht.Cell(r, 18)
		//PsComment, _ := Sht.Cell(r, 19)
		//PsCreaDate, _ := Sht.Cell(r, 20)
		//PsMajDate, _ := Sht.Cell(r, 21)
		//PsMajSrc, _ := Sht.Cell(r, 22)
		//PsAbdDate, _ := Sht.Cell(r, 23)
		//PsAbdSrc, _ := Sht.Cell(r, 24)

		_ = Wba.SetCellValue("Sheet1", fmt.Sprintf("A%v", r+1), val[0])

		//PsCode.Value = val[0]
		//PsNum.Value = val[1]
		//Ps1.Value = val[2]
		//FoNumTube1.Value = fo1[0]
		//FoColor1.Value = fo1[1]
		//CbEti1.Value = fo1[2]
		//Ps2.Value = val[3]
		//FoNumTube2.Value = fo2[0]
		//FoColor2.Value = fo2[1]
		//CbEti2.Value = fo2[2]
		//PsCsCode.Value = val[4]
		//CsNum.Value = cs[0]
		//BpEti.Value = bp
		//PsTiCode.Value = val[5]
		//TiEti.Value = ti
		//PsType.Value = val[6]
		//PsFunc.Value = val[7]
		//PsState.Value = val[8]
		//PsPreaff.Value = val[9]
		//PsComment.Value = val[10]
		//PsCreaDate.Value = val[11]
		//PsMajDate.Value = val[12]
		//PsMajSrc.Value = val[13]
		//PsAbdDate.Value = val[14]
		//PsAbdSrc.Value = val[15]

		if len(TPositionErr) > 0 {
			if checkPosErr(val[0]) {
				style := xlsx.NewStyle()
				style.Font.Color = "FFC000"
				//PsCode.SetStyle(style)
			}
		}

		NbrTot++
		loger.Void(fmt.Sprintf("%v/%v", NbrTot, d.NbrPos))
	}

	loger.Ok(fmt.Sprintf("%v positions concaténées", NbrTot))
}

func setFormatingWb() {
	headers := map[string]string{
		"A1": "ps_code",
		"B1": "ps_numero",
		"C1": "ps_1",
		"D1": "fo_numtub",
		"E1": "fo_color",
		"F1": "cb_etiquet",
		"G1": "ps_2",
		"H1": "fo_numtub",
		"I1": "fo_color",
		"J1": "cb_etiquet",
		"K1": "ps_cs_code",
		"L1": "cs_num",
		"M1": "bp_etiquet",
		"N1": "ps_ti_code",
		"O1": "ti_etiquet",
		"P1": "ps_type",
		"Q1": "ps_fonct",
		"R1": "ps_etat",
		"S1": "ps_preaff",
		"T1": "ps_comment",
		"U1": "ps_creadat",
		"V1": "ps_majdate",
		"W1": "ps_majsrc",
		"X1": "ps_abddate",
		"Y1": "ps_abdsrc",
	}

	for header := range headers {
		_ = Wba.SetCellValue("Sheet1", header, headers[header])
	}

	if err := Wba.SetPageLayout("Sheet1", &fitToWidth); err != nil {
		fmt.Println(err)
	}

	_ = Wba.SetColWidth("Sheet1", "A", "Y", 18)
	_ = Wba.SetColWidth("Sheet1", "B", "B", 10)
	_ = Wba.SetColWidth("Sheet1", "D", "E", 10)
	_ = Wba.SetColWidth("Sheet1", "H", "I", 10)
	_ = Wba.SetColWidth("Sheet1", "L", "L", 10)
	_ = Wba.SetColWidth("Sheet1", "P", "R", 8)

	//Sht.SetColWidth(1, 25, 18)
	//Sht.SetColWidth(2, 2, 10)
	//Sht.SetColWidth(4, 5, 10)
	//Sht.SetColWidth(8, 9, 10)
	//Sht.SetColWidth(12, 12, 10)
	//Sht.SetColWidth(16, 18, 8)
}

func checkPosErr(ps string) bool {
	for _, data := range TPositionErr {
		if ps == data.PsCode {
			return true
		}
	}
	return false
}

func getDataFibre(ps string) []string {
	for _, data := range TFibre {
		if ps == data.FoCode {

			var cb string
			for _, cbs := range TCable {
				if cbs.CbCode == data.FoCbCode {
					cb = cbs.CbEti
					break
				}
			}

			return []string{data.FoNumTube, data.FoColor, cb}
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
func getDataTirroir(ti string) string {
	for _, data := range TTirroir {
		if ti == data.TiCode {
			return data.TiEti
		}
	}
	return ""
}

func (d *Data) getFolderDLG() string {
	var dlg string

	err := filepath.Walk(d.SrcFile, func(path string, fileInfo os.FileInfo, err error) error {
		if fileInfo.IsDir() && strings.Contains(fileInfo.Name(), "-DLG-") && !strings.Contains(fileInfo.Name(), "_") {
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
func (d *Data) getDLGErr() string {
	var dlg string

	err := filepath.Walk(d.SrcFile, func(path string, fileInfo os.FileInfo, err error) error {
		if !fileInfo.IsDir() && strings.Contains(fileInfo.Name(), "-DLG-") && strings.Contains(fileInfo.Name(), "ERREURS-") {
			dlg = fileInfo.Name()
			return nil
		}
		return nil
	})

	if err != nil {
		loger.Error("Error pendant le listing des fichiers", err)
	}

	return dlg
}
