package pkg

import (
	"ConcatFiles/loger"
	"bufio"
	"encoding/csv"
	"fmt"
	"github.com/tealeg/xlsx"
	"io"
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

func CopyDir(source string, dest string) error {
	// Ouvre le dossier source
	sourceInfo, err := os.Stat(source)
	if err != nil {
		return err
	}
	if !sourceInfo.IsDir() {
		return &SourceNotDirectoryError{source}
	}

	// Crée le dossier de destination s'il n'existe pas encore
	err = os.MkdirAll(dest, sourceInfo.Mode())
	if err != nil {
		return err
	}

	// Parcours les fichiers du dossier source
	directory, err := os.Open(source)
	if err != nil {
		return err
	}
	defer directory.Close()

	files, err := directory.Readdir(-1)
	if err != nil {
		return err
	}

	// Copie le contenu de chaque fichier vers le dossier de destination
	for _, file := range files {
		sourceFile := source + "/" + file.Name()
		destFile := dest + "/" + file.Name()

		if file.IsDir() {
			// Copie le sous-dossier récursivement
			err = CopyDir(sourceFile, destFile)
			if err != nil {
				return err
			}
		} else {
			// Copie le fichier
			sourceStream, err := os.Open(sourceFile)
			if err != nil {
				return err
			}
			defer sourceStream.Close()

			destStream, err := os.Create(destFile)
			if err != nil {
				return err
			}
			defer destStream.Close()

			_, err = io.Copy(destStream, sourceStream)
			if err != nil {
				return err
			}
		}
	}

	return nil
}

// Erreur pour signaler que la source n'est pas un dossier
type SourceNotDirectoryError struct {
	source string
}

func (e *SourceNotDirectoryError) Error() string {
	return "Source '" + e.source + "' is not a directory"
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
			FoNintub:  val[5],
			FoCbCode:  val[2],
		}
		TFibre = append(TFibre, Item)
	}
}

func appendPosition(file string) {
	Csv := ReadCSV(file)

	for _, val := range Csv {
		Item := Position{
			PsCode:     val[0],
			PsNum:      val[1],
			Ps1:        val[2],
			Ps2:        val[3],
			PsCsCode:   val[4],
			PsTiCode:   val[5],
			PsType:     val[6],
			PsFunc:     val[7],
			PsState:    val[8],
			PsPreaff:   val[9],
			PsComment:  val[10],
			PsCreaDate: val[11],
			PsMajDate:  val[12],
			PsMajSrc:   val[13],
			PsAbdDate:  val[14],
			PsAbdSrc:   val[15],
		}
		TPosition = append(TPosition, Item)
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

func appendGraceAll() {
	for _, pos := range TPosition {

		fo1 := GetDataFibre(pos.Ps1)
		fo2 := GetDataFibre(pos.Ps2)
		cs := GetDataCassette(pos.PsCsCode)
		bp := GetDataEbp(cs[1])
		ti := GetDataTirroir(pos.PsTiCode)

		orderInt := 0
		if len(bp) > 1 {

			pb := "1"
			if bp[0:3] == "BPE" {
				pb = "2"
			} else if bp[0:3] == "PBO" {
				pb = "3"
			} else {
				pb = "4"
			}

			concat := pb + bp[len(bp)-3:] + cs[0]

			a, err := strconv.Atoi(concat)
			if err != nil {
				a = 0
			}

			orderInt = a
		}

		PsNumInt, _ := strconv.Atoi(pos.PsNum)
		FoNumTube1Int, _ := strconv.Atoi(fo1[0])
		FoNintub1Int, _ := strconv.Atoi(fo1[1])
		FoNumTube2Int, _ := strconv.Atoi(fo2[0])
		FoNintub2Int, _ := strconv.Atoi(fo2[1])
		CsNumInt, _ := strconv.Atoi(cs[0])

		Item := GraceAll{
			PsCode:     pos.PsCode,
			PsNum:      PsNumInt,
			Ps1:        pos.Ps1,
			FoNumTube1: FoNumTube1Int,
			FoNintub1:  FoNintub1Int,
			CbEti1:     fo1[2],
			Ps2:        pos.Ps2,
			FoNumTube2: FoNumTube2Int,
			FoNintub2:  FoNintub2Int,
			CbEti2:     fo2[2],
			PsCsCode:   pos.PsCsCode,
			CsNum:      CsNumInt,
			BpEti:      bp,
			PsTiCode:   pos.PsTiCode,
			TiEti:      ti,
			PsType:     pos.PsType,
			PsFunc:     pos.PsFunc,
			PsState:    pos.PsState,
			PsPreaff:   pos.PsPreaff,
			PsComment:  pos.PsComment,
			PsCreaDate: pos.PsCreaDate,
			PsMajDate:  pos.PsMajDate,
			PsMajSrc:   pos.PsMajSrc,
			PsAbdDate:  pos.PsAbdDate,
			PsAbdSrc:   pos.PsAbdSrc,
			OrderInt:   orderInt,
		}
		TGraceAll = append(TGraceAll, Item)
	}
}

func appendGraceLight() {
	for _, pos := range TPosition {

		cs := GetDataCassette(pos.PsCsCode)
		bp := GetDataEbp(cs[1])

		orderInt := 0
		if len(bp) > 1 {

			pb := "1"
			if bp[0:3] == "BPE" {
				pb = "2"
			} else if bp[0:3] == "PBO" {
				pb = "3"
			} else {
				pb = "4"
			}

			concat := pb + bp[len(bp)-3:] + cs[0]

			a, err := strconv.Atoi(concat)
			if err != nil {
				a = 0
			}

			orderInt = a
		}

		PsNumInt, _ := strconv.Atoi(pos.PsNum)
		CsNumInt, _ := strconv.Atoi(cs[0])

		Item := GraceLight{
			PsCode:     pos.PsCode,
			PsNum:      PsNumInt,
			Ps1:        pos.Ps1,
			Ps2:        pos.Ps2,
			PsCsCode:   pos.PsCsCode,
			CsNum:      CsNumInt,
			BpEti:      bp,
			PsTiCode:   pos.PsTiCode,
			PsType:     pos.PsType,
			PsFunc:     pos.PsFunc,
			PsState:    pos.PsState,
			PsPreaff:   pos.PsPreaff,
			PsComment:  pos.PsComment,
			PsCreaDate: pos.PsCreaDate,
			PsMajDate:  pos.PsMajDate,
			PsMajSrc:   pos.PsMajSrc,
			PsAbdDate:  pos.PsAbdDate,
			PsAbdSrc:   pos.PsAbdSrc,
			OrderInt:   orderInt,
		}
		TGraceLight = append(TGraceLight, Item)
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

			return []string{data.FoNumTube, data.FoNintub, cb}
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
