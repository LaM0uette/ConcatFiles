package pkg

import (
	"ConcatFiles/loger"
	"bufio"
	"encoding/csv"
	"fmt"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"os"
	"path"
	"path/filepath"
	"strconv"
	"strings"
)

func GetCurrentDir() string {
	pwd, err := os.Getwd()
	if err != nil {
		loger.Error("Error with current dir:", err)
		os.Exit(1)
	}
	return pwd
}

func FileExist(file string) bool {

	if _, err := os.Stat(file); err == nil {
		return true
	} else {
		return false
	}
}

func CreateNewFolder(path string) {
	if err := os.MkdirAll(path, os.ModePerm); err != nil {
		loger.Error("Erreur durant la création du dossier", err)
	}
}

func CopyFile(srcFile, dstFile string) {
	input, err := ioutil.ReadFile(srcFile)
	if err != nil {
		loger.Error("Erreur durant la copie du fichier (copier)", err)
	}

	err = ioutil.WriteFile(dstFile, input, 0644)
	if err != nil {
		loger.Error("Erreur durant la copie du fichier (coller)", err)
	}
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

func (d *Data) getFolderDLG() string {
	var dlg string

	if len(os.Args) > 2 {
		DrawParam("DOSSIER COURANT:", "OK")
		return filepath.Base(os.Args[2])
	}

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
		if !fileInfo.IsDir() && strings.Contains(fileInfo.Name(), "-DLG-") && strings.Contains(fileInfo.Name(), "ERREURS-") && strings.Contains(fileInfo.Name(), ".xlsx") {
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

func (d *Data) checkIfErrExist() {
	dlg := d.getDLGErr()

	if len(dlg) <= 0 {
		return
	}

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

			DrawParam("NOMBRES D'ERREURS:", strconv.Itoa(counter))
			break
		}

		if sheet.Name == "EBP" {
			counter := 0

			for i := 0; i < sheet.MaxRow; i++ {
				cell, _ := sheet.Cell(i, 1)
				if strings.Contains(cell.Value, "BP") {
					bp := EbpErr{BpCode: cell.Value}
					TEbpErr = append(TEbpErr, bp)
					counter++
				}
			}

			DrawParam("NOMBRES D'ERREURS:", strconv.Itoa(counter))
			break
		}
	}

}

//...
// Append
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

func appendPtech(file string) {
	Csv := ReadCSV(file)

	for _, val := range Csv {
		Item := Ptech{
			PtCode:   val[0],
			PtEti:    val[2],
			PtNdCode: val[3],
			PtAdCode: val[4],
		}
		TPtech = append(TPtech, Item)
	}
}

func appendReference(file string) {
	Csv := ReadCSV(file)

	for _, val := range Csv {
		Item := Reference{
			RfCode:   val[0],
			RfType:   val[1],
			RfFabric: val[2],
		}
		TReference = append(TReference, Item)
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

//...
// GetData
func GetDataCassette(cs string) []string {
	for _, data := range TCassette {
		if cs == data.CsCode {
			return []string{data.CsNum, data.CsBpCode}
		}
	}
	return []string{"", ""}
}

func GetDataEbp(bp string) string {
	for _, data := range TEbp {
		if bp == data.BpCode {
			return data.BpEti
		}
	}
	return ""
}

func GetDataFibre(ps string) []string {
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

func GetDataPtech(pt string) []string {
	for _, data := range TPtech {
		if pt == data.PtCode {
			return []string{data.PtEti, data.PtNdCode, data.PtAdCode}
		}
	}
	return []string{"", "", ""}
}

func GetDataReference(rf string) []string {
	for _, data := range TReference {
		if rf == data.RfCode {
			return []string{data.RfType, data.RfFabric}
		}
	}
	return []string{"", ""}
}

func GetDataTirroir(ti string) string {
	for _, data := range TTirroir {
		if ti == data.TiCode {
			return data.TiEti
		}
	}
	return ""
}
