package structure

type ExtRst struct {
	Reserved [2]byte
	Cb       [2]byte
	Phs      Phs
	Rphssub  RPHSSub
	Rgphruns []PhRuns
}
