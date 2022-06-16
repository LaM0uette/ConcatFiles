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

		pt := GetDataPtech(val[3])
		rf := GetDataReference(val[17])

		_ = Wba.SetCellValue(Sht, fmt.Sprintf("A%v", r), val[0])   //bp_code
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("B%v", r), val[1])   //bp_etiquet
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("C%v", r), val[2])   //bp_codeext
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("D%v", r), val[3])   //t_ptech
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("E%v", r), pt[0])    //pt_etiquet
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("F%v", r), pt[1])    //pt_nd_code
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("G%v", r), pt[2])    //pt_ad_code
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("H%v", r), val[4])   //bp_lt_code
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("I%v", r), val[5])   //bp_sf_code
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("J%v", r), val[6])   //bp_prop
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("K%v", r), val[7])   //bp_gest
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("L%v", r), val[8])   //bp_user
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("M%v", r), val[9])   //bp_proptyp
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("N%v", r), val[10])  //bp_statut
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("O%v", r), val[11])  //bp_etat
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("P%v", r), val[12])  //bp_occp
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("Q%v", r), val[13])  //bp_datemes
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("R%v", r), val[14])  //bp_avct
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("S%v", r), val[15])  //bp_typephy
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("T%v", r), val[16])  //bp_typelog
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("U%v", r), val[17])  //t_ref
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("V%v", r), rf[0])    //rf_type
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("W%v", r), rf[1])    //rf_fabric
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("X%v", r), val[18])  //bp_entrees
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("Y%v", r), val[19])  //bp_ref_kit
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("Z%v", r), val[20])  //bp_ca_nb
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AA%v", r), val[21]) //bp_nb_pas
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AB%v", r), val[22]) //bp_linecod
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AC%v", r), val[23]) //bp_oc_code
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AD%v", r), val[24]) //bp_racco
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AE%v", r), val[25]) //bp_comment
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AF%v", r), val[26]) //bp_creadat
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AG%v", r), val[27]) //bp_majdate
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AH%v", r), val[28]) //bp_majsrc
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AI%v", r), val[29]) //bp_abddate
		_ = Wba.SetCellValue(Sht, fmt.Sprintf("AJ%v", r), val[30]) //bp_abdsrc

		if len(TEbpErr) > 0 {
			if checkEbpErr(val[0]) {
				style, _ := Wba.NewStyle(fmt.Sprintf("{\"fill\":{\"type\":\"pattern\",\"color\":[\"#%s\"],\"pattern\":1}}", "FFC000"))
				_ = Wba.SetCellStyle(Sht, fmt.Sprintf("A%v", r), fmt.Sprintf("A%v", r), style)
			}
		}

		NbrTot++
		loger.Void(fmt.Sprintf("%v/%v", NbrTot, d.NbrItems))
	}

	d.setFormatingWbEbp()
	loger.Ok(fmt.Sprintf("%v ebp concaténées", NbrTot))
}

//...
// Actions
func (d *Data) setFormatingWbEbp() {
	headers := map[string]string{
		"A1":  "bp_code",
		"B1":  "bp_etiquet",
		"C1":  "bp_codeext",
		"D1":  "t_ptech",
		"E1":  "pt_etiquet",
		"F1":  "pt_nd_code",
		"G1":  "pt_ad_code",
		"H1":  "bp_lt_code",
		"I1":  "bp_sf_code",
		"J1":  "bp_prop",
		"K1":  "bp_gest",
		"L1":  "bp_user",
		"M1":  "bp_proptyp",
		"N1":  "bp_statut",
		"O1":  "bp_etat",
		"P1":  "bp_occp",
		"Q1":  "bp_datemes",
		"R1":  "bp_avct",
		"S1":  "bp_typephy",
		"T1":  "bp_typelog",
		"U1":  "t_ref",
		"V1":  "rf_type",
		"W1":  "rf_fabric",
		"X1":  "bp_entrees",
		"Y1":  "bp_ref_kit",
		"Z1":  "bp_ca_nb",
		"AA1": "bp_nb_pas",
		"AB1": "bp_linecod",
		"AC1": "bp_oc_code",
		"AD1": "bp_racco",
		"AE1": "bp_comment",
		"AF1": "bp_creadat",
		"AG1": "bp_majdate",
		"AH1": "bp_majsrc",
		"AI1": "bp_abddate",
		"AJ1": "bp_abdsrc",
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

func checkEbpErr(bp string) bool {
	for _, data := range TEbpErr {
		if bp == data.BpCode {
			return true
		}
	}
	return false
}
