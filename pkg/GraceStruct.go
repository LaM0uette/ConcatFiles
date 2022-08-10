package pkg

import "github.com/qax-os/excelize"

type Cable struct {
	CbCode,
	CbEti string
}
type Cassette struct {
	CsCode,
	CsNum,
	CsBpCode string
}
type Ebp struct {
	BpCode,
	BpEti string
}
type Fibre struct {
	FoCode,
	FoNumTube,
	FoNintub,
	FoCbCode string
}
type Ptech struct {
	PtCode,
	PtEti,
	PtNdCode,
	PtAdCode string
}
type Reference struct {
	RfCode,
	RfType,
	RfFabric string
}
type Tirroir struct {
	TiCode,
	TiEti string
}

type PositionErr struct {
	PsCode string
}

type EbpErr struct {
	BpCode string
}

var (
	Wba        *excelize.File
	fitToWidth excelize.FitToWidth

	NameTCable     = "t_cable.csv"
	NameTCassette  = "t_cassette.csv"
	NameTEbp       = "t_ebp.csv"
	NameTFibre     = "t_fibre.csv"
	NameTPosition  = "t_position.csv"
	NameTPtech     = "t_ptech.csv"
	NameTReference = "t_reference.csv"
	NameTTiroir    = "t_tiroir.csv"

	TCable     []Cable
	TCassette  []Cassette
	TEbp       []Ebp
	TFibre     []Fibre
	TPtech     []Ptech
	TReference []Reference
	TTirroir   []Tirroir

	TPositionErr []PositionErr
	TEbpErr      []EbpErr
)
