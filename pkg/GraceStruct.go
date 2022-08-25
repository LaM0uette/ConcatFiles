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
type Position struct {
	PsCode,
	PsNum,
	Ps1,
	Ps2,
	PsCsCode,
	PsTiCode,
	PsType,
	PsFunc,
	PsState,
	PsPreaff,
	PsComment,
	PsCreaDate,
	PsMajDate,
	PsMajSrc,
	PsAbdDate,
	PsAbdSrc string
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

type GraceAll struct {
	PsCode     string
	PsNum      int
	Ps1        string
	FoNumTube1 int
	FoNintub1  int
	CbEti1,
	Ps2 string
	FoNumTube2 int
	FoNintub2  int
	CbEti2,
	PsCsCode string
	CsNum int
	BpEti,
	PsTiCode,
	TiEti,
	PsType,
	PsFunc,
	PsState,
	PsPreaff,
	PsComment,
	PsCreaDate,
	PsMajDate,
	PsMajSrc,
	PsAbdDate,
	PsAbdSrc string
	OrderInt int
}
type GraceLight struct {
	PsCode   string
	PsNum    int
	Ps1      string
	Ps2      string
	PsCsCode string
	CsNum    int
	BpEti,
	PsTiCode,
	PsType,
	PsFunc,
	PsState,
	PsPreaff,
	PsComment,
	PsCreaDate,
	PsMajDate,
	PsMajSrc,
	PsAbdDate,
	PsAbdSrc string
	OrderInt int
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
	TPosition  []Position
	TPtech     []Ptech
	TReference []Reference
	TTirroir   []Tirroir

	TPositionErr []PositionErr
	TEbpErr      []EbpErr

	TGraceAll   []GraceAll
	TGraceLight []GraceLight
)
