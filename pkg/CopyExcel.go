package pkg

func (d *Data) CopyExcelFile(file string) {
	CopyFile(file, d.XlFile)
	DrawParam("GENERATION DE LA FICHE D'EXPORT:", "OK")
}
