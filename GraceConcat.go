package GraceConcat

import "flag"

func main() {
	FlagMode := flag.String("m", "c", "Choisit le mode de concatenation")
	flag.Parse()

}
