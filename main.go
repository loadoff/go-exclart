package main

import (
	"flag"
	"fmt"
	"image"
	_ "image/jpeg"
	_ "image/png"
	"os"

	"github.com/loadoff/excl"
	pb "gopkg.in/cheggaaa/pb.v1"
)

func main() {
	var excelPath string
	var imagePath string
	var outputPath string
	var sheetName string
	var rowNo int
	var colNo int
	var err error
	flag.StringVar(&imagePath, "i", "", "image file path")
	flag.StringVar(&excelPath, "f", "", "input excel file path")
	flag.StringVar(&outputPath, "o", "./output.xlsx", "output excel file path")
	flag.StringVar(&sheetName, "sheet", "Sheet1", "output workbook sheet name")
	flag.IntVar(&rowNo, "row", 1, "start row no")
	flag.IntVar(&colNo, "col", 1, "start col no")
	flag.Usage = func() {
		fmt.Fprint(os.Stderr, `Usage of go-exclart
	-i IMAGEFILE: image file path.
	-f EXCELFILE: input excel book path.
	-o EXCELFILE: output excel book path.
	-sheet SHEET: output sheet name in excel book. 
	-row ROWNO: output row no.
	-col COLNO: output col no.
`)
	}
	flag.Parse()
	file, err := os.Open(imagePath)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer file.Close()
	img, _, err := image.Decode(file)
	if err != nil {
		fmt.Println(err)
		return
	}
	rect := img.Bounds()
	var book *excl.Workbook
	if excelPath == "" {
		book, _ = excl.Create()
	} else {
		book, err = excl.Open(excelPath)
		if err != nil {
			fmt.Println(err)
			return
		}
	}
	defer book.Close()
	sheet, err := book.OpenSheet(sheetName)
	if err != nil {
		fmt.Println(err)
		return
	}
	sheet.ShowGridlines(false)
	for i := 0; i < rect.Max.X; i++ {
		sheet.SetColWidth(0.125, i+colNo)
	}
	fmt.Println("Drawing start!")
	bar := pb.StartNew(rect.Max.Y)
	for i := 0; i < rect.Max.Y; i++ {
		row := sheet.GetRow(i + rowNo)
		row.SetHeight(0.75)
		for j := 0; j < rect.Max.X; j++ {
			r, g, b, a := img.At(j, i).RGBA()
			r = r >> 8
			g = g >> 8
			b = b >> 8
			a = a >> 8
			if a != 0 {
				color := fmt.Sprintf("%02x%02x%02x%02x", a, r, g, b)
				cell := row.GetCell(j + colNo)
				cell.SetBackgroundColor(color)
			}
		}
		if i > 0 && i%100 == 0 {
			sheet.OutputThroughRowNo(i + rowNo)
		}
		bar.Increment()
	}
	bar.FinishPrint("Drawing end!")

	sheet.Close()

	fmt.Println("Create Excel book.")
	book.Save(outputPath)
	fmt.Println("Finished!")
}
