package excel2pdf

import (
	"testing"
	"log"
	"github.com/xuri/excelize/v2"
	"github.com/signintech/gopdf"
)

func TestExcel2Pdf(t *testing.T){
	log.Println("TestExcel2Pdf ...")
	
	//创建excel对象
	fe,err:=excelize.OpenFile("test.xlsx")
	if err!=nil {
		log.Println(err)
		return
	}
	//创建pdf对象
	pdf := &gopdf.GoPdf{}
	pdf.Start(gopdf.Config{PageSize: *gopdf.PageSizeA4})
	//加载字体
	pdf.AddTTFFont("仿宋", "./font/仿宋.ttf")
	//设置默认字体
	pdf.SetFont("仿宋", "", 12)
	
	err=Excel2Pdf(fe,pdf)
	if err!=nil {
		log.Println(err)
		t.Fail()
		return
	}

	err=pdf.WritePdf("test.pdf")
	if err!=nil {
		log.Println(err)
		t.Fail()
		return
	}
}