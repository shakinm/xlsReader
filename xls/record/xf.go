package record

import "github.com/shakinm/xlsReader/helpers"

//XF: Extended Format (E0h)

var XFRecord = []byte{0xE0, 0x00} //(E0h)

/*

The XF record stores formatting properties. There are two different XF records, one
for cell records and another for style records. The fStyle bit is true if the XF is a
style XF . The ixfe of a cell record ( BLANK , LABEL , NUMBER , RK , and so on) points
to a cell XF record, and the ixfe of a STYLE record points to a style XF record.
Note: in previous BIFF versions, the record number for the XF record was 43h.
Prior to BIFF5, all number format information was included in FORMAT records in the
BIFF file. Beginning with BIFF5, many of the built-in number formats were moved to
an internal table and are no longer saved with the file as FORMAT records. Use the
ifmt to associate the built-in number formats with an XF record. However, the
internal number formats are no longer visible in the BIFF file.
The following table lists all the number formats that are now maintained internally.
Note: 17h through 24h are reserved for international versions and are
undocumented at this time.

Index to internal
format (ifmt)			Format string
-------------------------------------
00h 					General
01h 					0
02h 					0.00
03h 					#,##0
04h 					#,##0.00
05h 					($#,##0_);($#,##0)
06h 					($#,##0_);[Red]($#,##0)
07h 					($#,##0.00_);($#,##0.00)
08h 					($#,##0.00_);[Red]($#,##0.00)
09h 					0%
0ah 					0.00%
0bh 					0.00E+00
0ch 					# ?/?
0dh 					# ??/??
0eh 					m/d/yy
0fh 					d-mmm-yy
10h 					d-mmm
11h 					mmm-yy
12h 					h:mm AM/PM
13h 					h:mm:ss AM/PM
14h 					h:mm
15h 					h:mm:ss
16h 					m/d/yy h:mm
25h 					(#,##0_);(#,##0)
26h 					(#,##0_);[Red](#,##0)
27h 					(#,##0.00_);(#,##0.00)
28h 					(#,##0.00_);[Red](#,##0.00)
29h 					_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
2ah 					_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)
2bh 					_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
2ch 					_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)
2dh 					mm:ss
2eh 					[h]:mm:ss
2fh 					mm:ss.0
30h 					##0.0E+0
31h 					@

A BIFF file can contain as many XF records as are necessary to describe the different
cell formats and styles in a workbook. The XF records are written in a table in the
workbook ( Book ) stream, and the index to the XF record table is a 0-based number
called ixfe .
The workbook stream must contain a minimum XF table consisting of 15 style XF
records and one cell XF record ( ixfe=0 through ixfe=15 ). The first XF record
( ixfe=0 ) is the XF record for the Normal style. The next 14 records ( ixfe=1
through ixfe=14 ) are XF records that correspond to outline styles RowLevel_1,
ColLevel_1, RowLevel_2, ColLevel_2, and so on. The last record ( ixfe=15 ) is the
default cell XF for the workbook.
Following these XF records are five additional style XF records (not strictly required)
that correspond to the Comma, Comma [0], Currency, Currency [0], and Percent
styles.


Cell XF Record — BIFF8
Record Data
Offset		Bits		Mask		Name		Contents
--------------------------------------------------------
4 			15–0 		FFFFh 		ifnt 		Index to the FONT record.
6 			15–0 		FFFFh 		ifmt 		Index to the FORMAT record.
8 			0 			0001h 		fLocked 	=1 if the cell is locked
			1 			0002h 		fHidden 	=1 if the cell is hidden.
			2 			0004h 		fStyle 		=0 for cell XF.
												=1 for style XF.
			3 			0008h 		f123Prefix	If the Transition Navigation Keys option is off (Options dialog box,
												Transition tab), f123Prefix=1 indicates that a leading apostrophe
												(single quotation mark) is being used to coerce the cell‘s contents to a
												simple string. If the Transition Navigation Keys option is on, f123Prefix=1 indicates
												that the cell formula begins with one of the four Lotus 1-2-3 alignment
												prefix characters:
																	' left
																	" right
																	^ centered
																	\ fill
												This bit is always 0 if fStyle=1 .
			15–4 		FFF0h		ixfParent	Index to the XF record of the parent style. Every cell XF must have a
												parent style XF , which is usually ixfeNormal=0 T his structure is always FFFh if fStyle=1 .
10			2–0			0007h		alc			Alignment:
													0= general
													1= left
													2= center
													3= right
													4= fill
													5= justify
													6= center across selection
			3			0008h		fWrap		=1 wrap text in cell.
			6–4			0070h		alcV		Vertical alignment:
													0= top
													1= center
													2= bottom
													3= justify
			7			0080h		fJustLast		(Used only in East Asian versions of Excel).
			15–8		FF00h		trot			Rotation, in degrees; 0–90dec is up  0–90 deg., 91–180dec is down 1–90
													deg, and 255dec is vertical.
12			3–0			000Fh		cIndent			Indent value (Format Cells dialog box, Alignment tab)
			4			0010h		fShrinkToFit	=1 if Shrink To Fit option is on
			5			0020h		fMergeCell		=1 if Merge Cells option is on (Format Cells dialog box, Alignment tab).
			7–6			00C0h		iReadOrder		Reading direction (East Asian versions only):
														0= Context
														1= Left-to-right
														2= Right-to-left
			9–8			0300h		(Reserved)
			10			0400h		fAtrNum			=1 if the ifmt is not equal to the ifmt of the parent style XF .
													This bit is N/A if fStyle=1 .
			11			0800h		fAtrFnt			=1 if the ifnt is not equal to the ifnt of the parent style XF .
													This bit is N/A if fStyle=1 .
			12			1000h		fAtrAlc			=1 if either the alc or the fWrap structure is not equal to the corresponding structure
													of the parent style XF . This bit is N/A if fStyle=1 .
			13			2000h		fAtrBdr			=1 if any border line structure ( dgTop , and so on) is not equal to the
													corresponding structure of the parent style XF.  This bit is N/A if fStyle=1 .
			14			4000h		fAtrPat			=1 if any pattern structure ( fls , icvFore , icvBack ) is not equal to
													the corresponding structure of the parent style XF . This bit is N/A if fStyle=1 .
			15			8000h		fAtrProt		=1 if either the fLocked structure or the fHidden structure is not equal to the
													corresponding structure of the parent style XF. This bit is N/A if fStyle=1.
14			3–0			000Fh		dgLeft			Border line style (see the following table).
			7–4			00F0h		dgRight			Border line style (see the following table).
			11–8		0F00h		dgTop			Border line style (see the following table).
			15–12		F000h		dgBottom		Border line style (see the following table).
16			6–0			007Fh		icvLeft			Index to the color palette for the left border color.
			13–7		3F80h		icvRight		Index to the color palette for the right border color.
			15–14		C000h		grbitDiag		1=diag down, 2=diag up, 3=both.
18			6–0			0000007Fh	icvTop			Index to the color palette for the top border color.
			13–7		00003F80h	icvBottom		Index to the color palette for the bottom border color.
			20–14		001FC000h	icvDiag			for diagonal borders.
			24–21		01E00000h	dgDiag			Border line style (see the following table).
			25			02000000h	fHasXFExt		=1 when a subsequent XFEXT record may modify the properties of this XF.
													New for Office Excel 2007
			31–26		FC000000h	fls				Fill pattern.
22			6–0			007Fh		icvFore			Index to the color palette for the foreground color of the fill pattern.
			13–7		3F80h		icvBack			Index to the color palette for the background color of the fill pattern.
			14			4000h		fSxButton		=1 if the XF record is attached to a PivotTable button. This bit is always 0 if fStyle=1 .
			15			8000h		(Reserved)

*/

type XF struct {
	font   [2]byte
	format [2]byte
	ttype  [2]byte
}

func (r *XF) Read(stream []byte) {
	copy(r.font[:], stream[0:2])
	copy(r.format[:], stream[2:4])
	copy(r.ttype[:], stream[4:6])

}

func (r *XF) GetFormatIndex() int {
	return int(helpers.BytesToUint16(r.format[:]))
}
