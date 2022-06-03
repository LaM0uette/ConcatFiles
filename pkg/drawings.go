package pkg

import (
	"ConcatFile/config"
	"ConcatFile/loger"
	"ConcatFile/rgb"
	"fmt"
	"time"
)

const (
	start = `
		 ██████╗ ██████╗ ███╗   ██╗ ██████╗ █████╗ ████████╗███████╗██╗██╗     ███████╗
		██╔════╝██╔═══██╗████╗  ██║██╔════╝██╔══██╗╚══██╔══╝██╔════╝██║██║     ██╔════╝
		██║     ██║   ██║██╔██╗ ██║██║     ███████║   ██║   █████╗  ██║██║     █████╗  
		██║     ██║   ██║██║╚██╗██║██║     ██╔══██║   ██║   ██╔══╝  ██║██║     ██╔══╝  
		╚██████╗╚██████╔╝██║ ╚████║╚██████╗██║  ██║   ██║   ██║     ██║███████╗███████╗
		 ╚═════╝ ╚═════╝ ╚═╝  ╚═══╝ ╚═════╝╚═╝  ╚═╝   ╚═╝   ╚═╝     ╚═╝╚══════╝╚══════╝`
	ligneSep = `■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■`

	author  = `Auteur:  `
	version = `Version: `
)

func DrawStart() {
	defer time.Sleep(1 * time.Second)

	loger.Ui(start)
	loger.Ui("\t\t", author+config.Author, "\n", "\t\t", version+config.Version)
	loger.Ui("\n")

	rgb.Green.Println(start)
	fmt.Print("\t\t", author+rgb.Green.Sprint(config.Author), "\n", "\t\t", version+rgb.Green.Sprint(config.Version))
	fmt.Print("\n\n")
}
