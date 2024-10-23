/*
Package excel2pdf use excelize lib to read excel and gofpdf lib to write pdf.
one sheet in excel will be one page in pdf.
*/
package excel2pdf

import (
	"github.com/signintech/gopdf"
	"github.com/xuri/excelize/v2"
	"fmt"
	"strings"
	"math"
)

const (
	BorderTypeLeft="left"
	BorderTypeRight="right"
	BorderTypeTop="top"
	BorderTypeBottom="bottom"
)

func Excel2Pdf(fe *excelize.File,pdf *gopdf.GoPdf)(error){
	//遍历所有的sheet
	for _, sheetName := range fe.GetSheetList() {
		pdf.AddPage()
		//处理单个sheet
		err := ProcessSheet(fe, sheetName,pdf)
		if err != nil {
			return err
		}
	}

	return nil
}

func ProcessSheet(fe *excelize.File, sheetName string,pdf *gopdf.GoPdf) error {
	dim,_:=fe.GetSheetDimension(sheetName)
	//基于冒号拆分dimension
	dimArr:=strings.Split(dim,":")
	maxCol,maxRow,_:=excelize.CellNameToCoordinates(dimArr[1])
	//计算每个cell的起始坐标位置
	rowPos,colPos:=GetCellPos(fe,sheetName,maxCol,maxRow)

	DrawSheetBorder(fe,sheetName,maxCol,maxRow,rowPos,colPos,pdf)

	DrawSheetContent(fe,sheetName,maxCol,maxRow,rowPos,colPos,pdf)

	return nil
}

func GetCellPos(fe *excelize.File, sheetName string,maxCol,maxRow int) (*[]float64,*[]float64) {
	//计算每个cell的起始坐标位置
	rowPos:=make([]float64,maxRow+1)
	
	rowPos[0]=0
	for row:=1;row<=maxRow;row++{
		rowHeight,_:=fe.GetRowHeight(sheetName,row)
		rowHeight=ConvertRowHeightToPixels(rowHeight)
		rowPos[row]=rowPos[row-1]+rowHeight
	}

	colPos:=make([]float64,maxCol+1)
	colPos[0]=0
	for col:=1;col<=maxCol;col++{
		colName,_:=excelize.ColumnNumberToName(col)
		colWidth,_:=fe.GetColWidth(sheetName,colName)
		colWidth=ConvertColWidthToPixels(colWidth)
		colPos[col]=colPos[col-1]+colWidth
	}

	return &rowPos,&colPos
}

func ConvertColWidthToPixels(width float64) float64 {
	var padding float64 = 5
	var pixels float64
	var maxDigitWidth float64 = 7
	if width == 0 {
		return pixels
	}
	if width < 1 {
		pixels = (width * 12) + 0.5
		return math.Ceil(pixels)
	}
	pixels = (width*maxDigitWidth + 0.5) + padding
	return math.Ceil(pixels)
}

func ConvertRowHeightToPixels(height float64) float64 {
	if height == 0 {
		return 0
	}
	return math.Ceil(4.0 / 3.4 * height)
}

func DrawSheetBorder(fe *excelize.File, sheetName string,maxCol,maxRow int,rowPos, colPos *[]float64,pdf *gopdf.GoPdf) {
	lineMap:=make(map[string]bool)

	for row:=1;row<=maxRow;row++{
		for col:=1;col<=maxCol;col++{
			cellName,_:=excelize.CoordinatesToCellName(col,row)
			styleIdx,_:=fe.GetCellStyle(sheetName,cellName)
			style,_:=fe.GetStyle(styleIdx)
			for _,border:=range style.Border {
				DrawBorder(row,col,&border,rowPos,colPos,pdf,&lineMap)
			}
		}
	}
}

func DrawBorder(row,col int,border *excelize.Border,rowPos, colPos *[]float64,pdf *gopdf.GoPdf,lineMap *map[string]bool) {
	var x1,y1,x2,y2 float64

	switch border.Type {
	case BorderTypeLeft:
		x1=(*colPos)[col-1]
		x2=(*colPos)[col-1]
		y1=(*rowPos)[row-1]
		y2=(*rowPos)[row]
	case BorderTypeRight:
		x1=(*colPos)[col]
		x2=(*colPos)[col]
		y1=(*rowPos)[row-1]
		y2=(*rowPos)[row]
	case BorderTypeTop:
		x1=(*colPos)[col-1]
		x2=(*colPos)[col]
		y1=(*rowPos)[row-1]
		y2=(*rowPos)[row-1]
	case BorderTypeBottom:
		x1=(*colPos)[col-1]
		x2=(*colPos)[col]
		y1=(*rowPos)[row]
		y2=(*rowPos)[row]
	}

	width:=1.0
	switch border.Style {
	case 2:
		width=2.0
	case 3:
		width=3.0
	default:
	}

	lineKeyStr:=fmt.Sprintf("%f_%f_%f_%f",x1,y1,x2,y2)
	if _,ok:=(*lineMap)[lineKeyStr];ok{
		return
	}

	(*lineMap)[lineKeyStr]=true

	pe:=PdfElement{
		Type:ElementTypeLine,
		Rect:[4]float64{x1,y1,x2,y2},
		Width:width,
	}

	DrawPdfElement(pdf,&pe)
}

func DrawSheetContent(fe *excelize.File, sheetName string,maxCol,maxRow int,rowPos, colPos *[]float64,pdf *gopdf.GoPdf) {
	mergedCells,_:=fe.GetMergeCells(sheetName)
	
	for row:=1;row<=maxRow;row++{
		for col:=1;col<=maxCol;col++ {
			//如果是合并单元格，则跳过
			if IsMergedCell(mergedCells,row,col)==true{
				continue
			}

			endRow,endCol:=GetEndAxis(mergedCells,row,col)

			cellName,_:=excelize.CoordinatesToCellName(col,row)
			cellValue,_:=fe.GetCellValue(sheetName,cellName)
			if cellValue!=""{
				styleIdx,_:=fe.GetCellStyle(sheetName,cellName)
				style,_:=fe.GetStyle(styleIdx)
				DrawText(row-1,col-1,endRow,endCol,cellValue,rowPos,colPos,pdf,style)
			}
		}
	}
}

func GetEndAxis(mergedCells []excelize.MergeCell,row,col int) (int,int) {
	for _,mc:=range mergedCells{
		startCol,startRow,_:=excelize.CellNameToCoordinates(mc.GetStartAxis())
		endCol,endRow,_:=excelize.CellNameToCoordinates(mc.GetEndAxis())
		if row>=startRow && row<=endRow && col>=startCol && col<=endCol{
			return endRow,endCol
		}
	}

	return row,col
}

func IsMergedCell(mergedCells []excelize.MergeCell,row,col int) bool {
	for _,mc:=range mergedCells{
		startCol,startRow,_:=excelize.CellNameToCoordinates(mc.GetStartAxis())
		endCol,endRow,_:=excelize.CellNameToCoordinates(mc.GetEndAxis())
		//第一个单元格不作为合并单元格
		if row==startRow && col==startCol {
			return false
		}

		if row>=startRow && row<=endRow && col>=startCol && col<=endCol{
			return true
		}
	}

	return false
}

func DrawText(startRow,startCol,endRow,endCol int,text string,rowPos, colPos *[]float64,pdf *gopdf.GoPdf,style *excelize.Style) {
	var x1,y1,x2,y2 float64

	x1=(*colPos)[startCol]
	x2=(*colPos)[endCol]
	y1=(*rowPos)[startRow]
	y2=(*rowPos)[endRow]

	fontSize:=style.Font.Size
	fontSize=ConvertRowHeightToPixels(fontSize)

	TextFont := &TextFont{
		Family: style.Font.Family,
		Size: fontSize,
	}

	pe:=PdfElement{
		Type:ElementTypeText,
		Rect:[4]float64{x1,y1,x2,y2},
		Content:text,
		Font:TextFont,
		LineHeight:fontSize+5,
		HorizontalAlign:style.Alignment.Horizontal,
		VerticalAlign:style.Alignment.Vertical,
		Padding:5,
		WordSpace:1,
	}

	DrawPdfElement(pdf,&pe)
}

func CreatePdfElement(pdf *gopdf.GoPdf, cell string) error {
	return nil
}