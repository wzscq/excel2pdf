package excel2pdf

import (
	"github.com/signintech/gopdf"
	"fmt"
	"strings"
)

const (
	ElementTypeLine="line"
	ElementTypeText="text"
)

const (
	ElementAlignLeft="left"
	ElementAlignCenter="center"
	ElementAlignRight="right"
)

const (
	VerticalAlignTop="top"
	VerticalAlignMiddle="center"
	VerticalAlignBottom="bottom"
)

type PdfLine struct {
	Width float64
	Height float64
	Content string
}

type ElementColor struct {
	R uint8 `json:"r,omitempty"`
	G uint8 `json:"g,omitempty"`
	B uint8 `json:"b,omitempty"`
}

type TextFont struct {
	Size float64 `json:"size,omitempty"`
	Style string `json:"style,omitempty"`
	Family string `json:"family,omitempty"`
}

type PdfElement struct {
	Type string `json:"type,omitempty"`
	Font *TextFont `json:"font,omitempty"`
	Color *ElementColor `json:"color,omitempty"`
	Rect [4]float64 `json:"rect,omitempty"`
	Content string `json:"content,omitempty"`
	Width float64 `json:"width,omitempty"`
	LineHeight float64  `json:"lineHeight,omitempty"`
	WordSpace float64 `json:"wordSpace,omitempty"`
	MaxContentLength int `json:"maxContentLength,omitempty"`
	VerticalAlign string `json:"verticalAlign,omitempty"`
	HorizontalAlign string `json:"horizontalAlign,omitempty"`
	Padding float64 `json:"padding,omitempty"`
}

func DrawPdfElement(pdf *gopdf.GoPdf,element *PdfElement){
	//创建pdf元素
	switch element.Type {
	case ElementTypeText:
		DrawPdfText(pdf,element)
	case ElementTypeLine:
		DrawPdfLine(pdf,element)
	default:
		fmt.Println("Unknow element type:",element.Type)
	}
}

func DrawPdfText(pdf *gopdf.GoPdf,element *PdfElement){
	//创建pdf文本元素
	//获取文本内容
	text:=element.Content

	if element.MaxContentLength>0 && len(text)>element.MaxContentLength {
		text=text[0:element.MaxContentLength]
	}

	//设置字体
	if element.Font!=nil {
		//log.Println("Font:",element.Font.Family,element.Font.Style,element.Font.Size)
		pdf.SetFont(element.Font.Family, element.Font.Style, element.Font.Size)
	}
	//设置颜色
	if element.Color!=nil {
		pdf.SetTextColor(element.Color.R,element.Color.G,element.Color.B)
	}

	xPos:=element.Rect[0]+element.Padding 
	yPos:=element.Rect[1]+element.Padding
	boxWidth := element.Rect[2] - element.Rect[0]-element.Padding*2
	boxHeight := element.Rect[3] - element.Rect[1]-element.Padding
	lineHeight:=element.LineHeight
	wordSpace:=element.WordSpace

	//首先对文本按照换行符进行拆分
	lines:=strings.Split(text,"\n")
	//计算所有文本行的宽度和高度，转换为pdfLine对象
	pdfLines:=GetPdfLines(pdf,lines,lineHeight,xPos,yPos,boxWidth,wordSpace)
	//计算垂直对齐后的yPos
	yPos=GetYPos(pdfLines,element.VerticalAlign,yPos,boxHeight)

	//输出文本
	for _,line:=range pdfLines {
		left:=GetXPos(&line,element.HorizontalAlign,xPos,boxWidth)
		runeSlice := []rune(line.Content)
		pdf.SetY(yPos)
		for i := 0; i < len(runeSlice); i++ {
			word := string(runeSlice[i]) + "" // 加一个空格，避免中文字符连续
			wordWidth, _ := pdf.MeasureTextWidth(word)
			pdf.SetX(left)
			pdf.Cell(nil, word)
			left=left+wordWidth+wordSpace
		}
		yPos=yPos+line.Height
	}
}

func GetXPos(line *PdfLine,horizontalAlign string,xPos,boxWidth float64)float64{
	//如果是左对齐，则直接返回原始xPos
	if horizontalAlign==ElementAlignLeft {
		return xPos
	}

	//如果文本宽度大于boxWidth，则返回原始xPos
	if line.Width>boxWidth {
		return xPos
	}

	//计算水平对齐后的xPos
	if horizontalAlign==ElementAlignCenter {
		return xPos+(boxWidth-line.Width)/2
	}
	
	return xPos+boxWidth-line.Width
}

func GetYPos(pdfLines []PdfLine,verticalAlign string,yPos,boxHeight float64)float64{
	//如果是顶部对齐，则直接返回原始yPos
	if verticalAlign==VerticalAlignTop {
		return yPos
	}
	
	//计算所有行的总高度
	totalHeight:=0.0
	for _,line:=range pdfLines {
		totalHeight=totalHeight+line.Height
	}

	//如果总高度大于boxHeight，则返回原始yPos
	if totalHeight>boxHeight {
		return yPos
	}

	//计算垂直对齐后的yPos
	if verticalAlign==VerticalAlignMiddle {
		return yPos+(boxHeight-totalHeight)/2
	}
	
	return yPos+boxHeight-totalHeight
}

func GetPdfLines(pdf *gopdf.GoPdf,lines []string,lineHeight,xPos,yPos,boxWidth,wordSpace float64)[]PdfLine{
	pdfLines:=[]PdfLine{}
	for _,line:=range lines {
		
		pefLine:=PdfLine{}
		runeSlice := []rune(line)
		lineWidth:=0.0
		
		for i := 0; i < len(runeSlice); i++ {
			word := string(runeSlice[i]) + "" // 加一个空格，避免中文字符连续
			wordWidth, _ := pdf.MeasureTextWidth(word)
			if lineWidth+wordWidth > boxWidth {
				pefLine.Width=lineWidth
				pefLine.Height=lineHeight
				pdfLines=append(pdfLines,pefLine)
				pefLine=PdfLine{}
				lineWidth=0.0
			}
			pefLine.Content=pefLine.Content+word
			lineWidth=lineWidth+wordWidth+wordSpace
		}

		pefLine.Width=lineWidth
		pefLine.Height=lineHeight
		pdfLines=append(pdfLines,pefLine)
	}
	return pdfLines
}

func DrawPdfLine(pdf *gopdf.GoPdf,element *PdfElement){
	//创建pdf线条元素
	//设置颜色
	if element.Color!=nil {
		pdf.SetStrokeColor(element.Color.R,element.Color.G,element.Color.B)
	}
	pdf.SetLineWidth(element.Width)
	//画线
	pdf.Line(element.Rect[0],element.Rect[1],element.Rect[2],element.Rect[3])
}