package record

/*
Record Data — BIFF8

Offset		Field Name	Size	Contents
--------------------------------------------
4 			vers 		2 		Version number:
								=0600 for BIFF8
6 			dt 			2 		Substream type:
									0005h = Workbook globals
									0006h = Visual Basic module
									0010h = Worksheet or dialog sheet
									0020h = Chart
									0040h = Excel 4.0 macro sheet
									0100h = Workspace file
8 			rupBuild 	2 		Build identifier (=0DBBh for Excel 97)
10 			rupYear 	2 		Build year (=07CCh for Excel 97)
12 			bfh 		4 		File history flags
16 			sfo 		4 		Lowest BIFF version (see text)


The rupBuild and rupYear fields contain numbers that identify the version (build)
 	of Excel that wrote the file. If you write a BIFF file, you can use the BiffView utility
	to determine the current values of these fields by examining a BOF record in a
	workbook file.
The sfo structure contains the earliest version ( vers structure) of Excel that can read all
	records in this file.

The bfh structure contains the following flag bits:

Bits 	Mask 		Flag Name 		Contents
--------------------------------------------
0 		00000001h 	fWin 			=1 if the file was last edited by Excel for Windows
1 		00000002h 	fRisc 			=1 if the file was last edited by Excel on a RISC platform
2 		00000004h 	fBeta 			=1 if the file was last edited by a beta version of Excel
3 		00000008h 	fWinAny 		=1 if the file has ever been edited by Excel for Windows
4 		00000010h 	fMacAny 		=1 if the file has ever been edited by Excel for the Macintosh
5 		00000020h 	fBetaAny 		=1 if the file has ever been edited by a beta version of Excel
7–6 	000000C0h 					(Reserved) Reserved; must be 0 (zero)
8		00000100h	fRiscAny		=1 if the file has ever been edited by Excel on a RISC platform
31–9 	FFFFFE00 					(Reserved) Reserved; must be 0 (zero)

*/
var FlagBIFF8 = []byte{0x00, 0x06}

type biff8 struct {
	vers     [2]byte
	dt       [2]byte
	rupBuild [2]byte
	rupYear  [2]byte
	bfh      [4]byte
	sfo      [4]byte
}

var FlagBIFF5 = []byte{0x00, 0x05}

type biff5 struct {
	vers     [2]byte
	dt       [2]byte
	rupBuild [2]byte
	rupYear  [2]byte
}
