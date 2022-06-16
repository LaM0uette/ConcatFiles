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
	d.runConcatEbp(path.Join(d.DstFile, NameTEbp))

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

func (d *Data) runConcatEbp(file string) {
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

		_ = Wba.SetCellValue(Sht, fmt.Sprintf("A%v", r), val[0])   //bp_code
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("B%v", r), val[1])   //bp_etiquet
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("C%v", r), val[2])   //bp_codeext
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("D%v", r), val[3])   //t_ptech
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("E%v", r), val[])   //pt_etiquet
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("F%v", r), val[])   //pt_nd_code
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("G%v", r), val[])   //pt_ad_code
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("H%v", r), val[4])   //bp_lt_code
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("I%v", r), val[5])   //bp_sf_code
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("J%v", r), val[6])   //bp_prop
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("K%v", r), val[7])   //bp_gest
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("L%v", r), val[8])    //bp_user
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("M%v", r), val[9])       //bp_proptyp
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("N%v", r), val[10])   //bp_statut
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("O%v", r), val[11])       //bp_etat
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("P%v", r), val[12])   //bp_occp
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("Q%v", r), val[13])   //bp_datemes
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("R%v", r), val[14])   //bp_avct
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("S%v", r), val[15])   //bp_typephy
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("T%v", r), val[16])  //bp_typelog
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("U%v", r), val[17])  //t_ref
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("V%v", r), val[])  //rf_code
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("W%v", r), val[])  //rf_type
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("X%v", r), val[])  //rf_fabric
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("Y%v", r), val[18])  //bp_entrees
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("Z%v", r), val[19])  //bp_ref_kit
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AA%v", r), val[20]) //bp_ca_nb
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AB%v", r), val[21]) //bp_nb_pas
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AC%v", r), val[22]) //bp_linecod
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AD%v", r), val[23]) //bp_oc_code
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AE%v", r), val[24]) //bp_racco
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AF%v", r), val[25]) //bp_comment
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AG%v", r), val[26]) //bp_creadat
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AH%v", r), val[27]) //bp_majdate
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AI%v", r), val[28]) //bp_majsrc
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AJ%v", r), val[29]) //bp_abddate
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AK%v", r), val[30]) //bp_abdsrc

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
	loger.Ok(fmt.Sprintf("%v ebp concaténées", NbrTot))
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
