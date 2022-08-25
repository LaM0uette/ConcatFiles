package pkg

import (
	"ConcatFiles/loger"
	"fmt"
	"github.com/qax-os/excelize"
	"path"
	"sort"
	"strconv"
)

func (d *Data) ConcatCSVGraceLight() {
	DrawSep("CONCAT CSV GRACE LIGHT")

	Wba, _ = excelize.OpenFile(d.XlFile)

	d.copyCSVGraceLight()
	DrawParam("COPIE DES CSV:", "OK")

	d.countPositionsLight()
	DrawParam("NOMBRE DE POSTIONS:", strconv.Itoa(d.NbrItems))

	d.checkIfErrExist()

	d.appendDatasInStructsLight()
	DrawParam("AJOUT DES DONNÉES DANS LES STRUCTS:", "OK")

	sortDataLight()
	DrawParam("TRIAGE DES DONNÉES:", "OK")

	DrawSep("COMPILATION")
	d.runConcatGraceLight()

	if err := Wba.SaveAs(d.XlFile); err != nil {
		loger.Error("Erreur pendant la sauvegarde du fichier Excel:", err)
	}
}

//...
// Functions
func (d *Data) copyCSVGraceLight() {
	dlgPath := path.Join(d.SrcFile, d.getFolderDLG())

	CopyFile(path.Join(dlgPath, NameTCable), path.Join(d.DstFile, NameTCable))
	CopyFile(path.Join(dlgPath, NameTCassette), path.Join(d.DstFile, NameTCassette))
	CopyFile(path.Join(dlgPath, NameTEbp), path.Join(d.DstFile, NameTEbp))
	CopyFile(path.Join(dlgPath, NameTFibre), path.Join(d.DstFile, NameTFibre))
	CopyFile(path.Join(dlgPath, NameTPosition), path.Join(d.DstFile, NameTPosition))
	CopyFile(path.Join(dlgPath, NameTTiroir), path.Join(d.DstFile, NameTTiroir))
}

func (d *Data) countPositionsLight() {
	tPositionPath := path.Join(d.DstFile, NameTPosition)
	CsvData := ReadCSV(tPositionPath)
	d.NbrItems = len(CsvData)
}

func sortDataLight() {
	sort.Slice(TGraceAll, func(i, j int) bool {
		return TGraceAll[i].OrderInt < TGraceAll[j].OrderInt
	})
}

func (d *Data) appendDatasInStructsLight() {
	appendPosition(path.Join(d.DstFile, NameTPosition))
	appendFibre(path.Join(d.DstFile, NameTFibre))
	appendCable(path.Join(d.DstFile, NameTCable))
	appendCassette(path.Join(d.DstFile, NameTCassette))
	appendEbp(path.Join(d.DstFile, NameTEbp))
	appendTirroir(path.Join(d.DstFile, NameTTiroir))

	appendGraceLight()
}

func (d *Data) runConcatGraceLight() {
	Sht := "Sheet1"
	NbrTot := 0

	for r, dt := range TGraceLight {
		r++

		_ = Wba.SetCellValue(Sht, fmt.Sprintf("A%v", r), dt.PsCode)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("B%v", r), dt.PsNum)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("C%v", r), dt.Ps1)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("D%v", r), dt.Ps2)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("E%v", r), dt.PsCsCode)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("F%v", r), dt.CsNum)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("G%v", r), dt.BpEti)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("H%v", r), dt.PsTiCode)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("I%v", r), dt.PsType)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("J%v", r), dt.PsFunc)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("K%v", r), dt.PsState)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("L%v", r), dt.PsPreaff)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("M%v", r), dt.PsComment)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("N%v", r), dt.PsCreaDate)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("O%v", r), dt.PsMajDate)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("P%v", r), dt.PsMajSrc)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("Q%v", r), dt.PsAbdDate)
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("R%v", r), dt.PsAbdSrc)

		if len(TPositionErr) > 0 {
			if checkPosErrLight(dt.PsCode) {
				style, _ := Wba.NewStyle(fmt.Sprintf("{\"fill\":{\"type\":\"pattern\",\"color\":[\"#%s\"],\"pattern\":1}}", "FFC000"))
				_ = Wba.SetCellStyle(Sht, fmt.Sprintf("A%v", r), fmt.Sprintf("A%v", r), style)
			}
		}

		NbrTot++
		loger.Void(fmt.Sprintf("%v/%v", NbrTot, d.NbrItems))
	}

	d.setFormatingWbGraceLight()
	loger.Ok(fmt.Sprintf("%v positions concaténées", NbrTot))
}

//...
// Actions
func (d *Data) setFormatingWbGraceLight() {
	headers := map[string]string{
		"A1": "ps_code",
		"B1": "ps_numero",
		"C1": "ps_1",
		"D1": "ps_2",
		"E1": "ps_cs_code",
		"F1": "cs_num",
		"G1": "bp_etiquet",
		"H1": "ps_ti_code",
		"I1": "ps_type",
		"J1": "ps_fonct",
		"K1": "ps_etat",
		"L1": "ps_preaff",
		"M1": "ps_comment",
		"N1": "ps_creadat",
		"O1": "ps_majdate",
		"P1": "ps_majsrc",
		"Q1": "ps_abddate",
		"R1": "ps_abdsrc",
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

func checkPosErrLight(ps string) bool {
	for _, data := range TPositionErr {
		if ps == data.PsCode {
			return true
		}
	}
	return false
}
