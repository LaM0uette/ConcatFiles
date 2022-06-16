package pkg

import (
	"ConcatFiles/loger"
	"fmt"
	"github.com/qax-os/excelize"
	"path"
	"strconv"
)

func (d *Data) ConcatCSVEbp() {
	DrawSep("CONCAT CSV EBP")

	Wba, _ = excelize.OpenFile(d.XlFile)

	d.copyCSVEbp()
	DrawParam("COPIE DES CSV:", "OK")

	d.countEbp()
	DrawParam("NOMBRE D'EBP:", strconv.Itoa(d.NbrItems))

	//d.checkIfErrExist()

	d.appendEbpDatasInStructs()
	DrawParam("AJOUT DES DONNÉES DANS LES STRUCTS:", "OK")

	DrawSep("COMPILATION")
	d.runConcat(path.Join(d.DstFile, NameTPosition))

	if err := Wba.SaveAs(d.XlFile); err != nil {
		loger.Error("Erreur pendant la sauvegarde du fichier Excel:", err)
	}
}

//...
// Functions
func (d *Data) copyCSVEbp() {
	dlgPath := path.Join(d.SrcFile, d.getFolderDLG())

	CopyFile(path.Join(dlgPath, NameTEbp), path.Join(d.DstFile, NameTEbp))
	CopyFile(path.Join(dlgPath, NameTPtech), path.Join(d.DstFile, NameTPtech))
	CopyFile(path.Join(dlgPath, NameTReference), path.Join(d.DstFile, NameTReference))
}

func (d *Data) countEbp() {
	tEbpPath := path.Join(d.DstFile, NameTEbp)
	CsvData := ReadCSV(tEbpPath)
	d.NbrItems = len(CsvData)
}

func (d *Data) appendEbpDatasInStructs() {
	appendPtech(path.Join(d.DstFile, NameTPtech))
	appendReference(path.Join(d.DstFile, NameTReference))
}

func (d *Data) runConcat(file string) {
	CsvData := ReadCSV(file)
	Sht := "Sheet1"
	NbrTot := 0

	for r, val := range CsvData {
		r++

		fo1 := getDataFibre(val[2])   //ps1 et cb1
		fo2 := getDataFibre(val[3])   //ps2 et cb2
		cs := getDataCassette(val[4]) //psCsCode
		bp := getDataEbp(cs[1])       //csCode
		ti := getDataTirroir(val[5])  //tiCode

		_ = Wba.SetCellValue(Sht, fmt.Sprintf("A%v", r), val[0])  //PsCode
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("B%v", r), val[1])  //PsNum
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("C%v", r), val[2])  //Ps1
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("D%v", r), fo1[0])  //FoNumTube1
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("E%v", r), fo1[1])  //FoColor1
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("F%v", r), fo1[2])  //CbEti1
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("G%v", r), val[3])  //Ps2
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("H%v", r), fo2[0])  //FoNumTube2
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("I%v", r), fo2[1])  //FoColor2
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("J%v", r), fo2[2])  //CbEti2
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("K%v", r), val[4])  //PsCsCode
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("L%v", r), cs[0])   //CsNum
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("M%v", r), bp)      //BpEti
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("N%v", r), val[5])  //PsTiCode
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("O%v", r), ti)      //TiEti
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("P%v", r), val[6])  //PsType
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("Q%v", r), val[7])  //PsFunc
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("R%v", r), val[8])  //PsState
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("S%v", r), val[9])  //PsPreaff
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("T%v", r), val[10]) //PsComment
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("U%v", r), val[11]) //PsCreaDate
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("V%v", r), val[12]) //PsMajDate
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("W%v", r), val[13]) //PsMajSrc
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("X%v", r), val[14]) //PsAbdDate
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("Y%v", r), val[15]) //PsAbdSrc

		if len(TPositionErr) > 0 {
			if checkPosErr(val[0]) {
				//style := xlsx.NewStyle()
				//style.Font.Color = "FFC000"

				style, _ := Wba.NewStyle(fmt.Sprintf("{\"fill\":{\"type\":\"pattern\",\"color\":[\"#%s\"],\"pattern\":1}}", "FFC000"))
				_ = Wba.SetCellStyle(Sht, fmt.Sprintf("A%v", r), fmt.Sprintf("A%v", r), style)
			}
		}

		NbrTot++
		loger.Void(fmt.Sprintf("%v/%v", NbrTot, d.NbrItems))
	}

	d.setFormatingWb()
	loger.Ok(fmt.Sprintf("%v positions concaténées", NbrTot))
}

//...
// Actions
func (d *Data) setFormatingWb() {
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

//...
// GetData
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
