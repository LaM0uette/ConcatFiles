package pkg

import (
	"ConcatFiles/loger"
	"fmt"
	"github.com/qax-os/excelize"
	"path"
	"strconv"
)

func (d *Data) ConcatCSVGrace() {
	DrawSep("CONCAT CSV GRACE")

	Wba, _ = excelize.OpenFile(d.XlFile)

	d.copyCSVGrace()
	DrawParam("COPIE DES CSV:", "OK")

	d.countPositions()
	DrawParam("NOMBRE DE POSTIONS:", strconv.Itoa(d.NbrItems))

	d.checkIfErrExist()

	d.appendDatasInStructs()
	DrawParam("AJOUT DES DONNÉES DANS LES STRUCTS:", "OK")

	sortPosition()
	DrawParam("TRIAGE DES DONNÉES:", "OK")

	DrawSep("COMPILATION")
	d.runConcatGrace(path.Join(d.DstFile, NameTPosition))

	if err := Wba.SaveAs(d.XlFile); err != nil {
		loger.Error("Erreur pendant la sauvegarde du fichier Excel:", err)
	}
}

//...
// Functions
func (d *Data) copyCSVGrace() {
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
	d.NbrItems = len(CsvData)
}

func sortPosition() {
	//sort.Slice(planets, func(i, j int) bool {
	//	return planets[i].Axis < planets[j].Axis
	//})
}

func (d *Data) appendDatasInStructs() {
	appendPosition(path.Join(d.DstFile, NameTPosition))
	appendFibre(path.Join(d.DstFile, NameTFibre))
	appendCable(path.Join(d.DstFile, NameTCable))
	appendCassette(path.Join(d.DstFile, NameTCassette))
	appendEbp(path.Join(d.DstFile, NameTEbp))
	appendTirroir(path.Join(d.DstFile, NameTTiroir))
}

func (d *Data) runConcatGrace(file string) {
	Sht := "Sheet1"
	NbrTot := 0

	for r, pos := range TPosition {
		r++

		fo1 := GetDataFibre(pos.Ps1)
		fo2 := GetDataFibre(pos.Ps2)
		cs := GetDataCassette(pos.PsCode)
		bp := GetDataEbp(cs[1])
		ti := GetDataTirroir(pos.PsTiCode)

		_ = Wba.SetCellValue(Sht, fmt.Sprintf("A%v", r), pos.PsCode)     //PsCode
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("B%v", r), pos.PsNum)      //PsNum
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("C%v", r), pos.Ps1)        //Ps1
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("D%v", r), fo1[0])         //FoNumTube1
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("E%v", r), fo1[1])         //FoColor1
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("F%v", r), fo1[2])         //CbEti1
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("G%v", r), pos.Ps2)        //Ps2
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("H%v", r), fo2[0])         //FoNumTube2
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("I%v", r), fo2[1])         //FoColor2
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("J%v", r), fo2[2])         //CbEti2
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("K%v", r), pos.PsCode)     //PsCsCode
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("L%v", r), cs[0])          //CsNum
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("M%v", r), bp)             //BpEti
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("N%v", r), pos.PsTiCode)   //PsTiCode
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("O%v", r), ti)             //TiEti
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("P%v", r), pos.PsType)     //PsType
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("Q%v", r), pos.PsFunc)     //PsFunc
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("R%v", r), pos.PsState)    //PsState
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("S%v", r), pos.PsPreaff)   //PsPreaff
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("T%v", r), pos.PsComment)  //PsComment
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("U%v", r), pos.PsCreaDate) //PsCreaDate
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("V%v", r), pos.PsMajDate)  //PsMajDate
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("W%v", r), pos.PsMajSrc)   //PsMajSrc
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("X%v", r), pos.PsAbdDate)  //PsAbdDate
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("Y%v", r), pos.PsAbdSrc)   //PsAbdSrc

		if len(TPositionErr) > 0 {
			if checkPosErr(pos.PsCode) {
				style, _ := Wba.NewStyle(fmt.Sprintf("{\"fill\":{\"type\":\"pattern\",\"color\":[\"#%s\"],\"pattern\":1}}", "FFC000"))
				_ = Wba.SetCellStyle(Sht, fmt.Sprintf("A%v", r), fmt.Sprintf("A%v", r), style)
			}
		}

		NbrTot++
		loger.Void(fmt.Sprintf("%v/%v", NbrTot, d.NbrItems))
	}

	d.setFormatingWbGrace()
	loger.Ok(fmt.Sprintf("%v positions concaténées", NbrTot))
}

//...
// Actions
func (d *Data) setFormatingWbGrace() {
	headers := map[string]string{
		"A1": "ps_code",
		"B1": "ps_numero",
		"C1": "ps_1",
		"D1": "fo_numtub",
		"E1": "fo_nintub",
		"F1": "cb_etiquet",
		"G1": "ps_2",
		"H1": "fo_numtub",
		"I1": "fo_nintub",
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

	_ = Wba.SetCellValue("MACRO", "A3", d.DstFile)
	Wba.SetActiveSheet(1)

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
