package rgb

import (
	"github.com/gookit/color"
)

var (
	Green   = color.NewRGBStyle(color.RGB(44, 168, 65))
	GreenB  = color.NewRGBStyle(color.RGB(44, 168, 65)).SetOpts(color.Opts{color.Bold})
	GreenBg = color.NewRGBStyle(color.RGB(240, 240, 240), color.RGB(44, 168, 65)).SetOpts(color.Opts{color.Bold})
	RedB    = color.NewRGBStyle(color.RGB(207, 25, 37)).SetOpts(color.Opts{color.Bold})
	RedBg   = color.NewRGBStyle(color.RGB(240, 240, 240), color.RGB(207, 25, 37)).SetOpts(color.Opts{color.Bold})
	//Yellow   = color.NewRGBStyle(color.RGB(196, 168, 27))
	YellowUB  = color.NewRGBStyle(color.RGB(196, 168, 27)).SetOpts(color.Opts{color.OpUnderscore, color.Bold})
	Majenta   = color.NewRGBStyle(color.RGB(204, 57, 199))
	MajentaBg = color.NewRGBStyle(color.RGB(240, 240, 240), color.RGB(204, 57, 199))

	Gray = color.NewRGBStyle(color.RGB(80, 80, 80))
)
