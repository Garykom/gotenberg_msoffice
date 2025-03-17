package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/gofiber/fiber/v2/log"
)

func officeWord2pdf(fileName string, pdfPath string) {
	log.Info("officeWord2pdf - start")
	log.Info("fileName=" + fileName)
	log.Info("pdfPath=" + pdfPath)

	ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("Word.Application")
	if err != nil {
		checkErr(err)
		return
	}

	defer unknown.Release()
	word, _ := unknown.QueryInterface(ole.IID_IDispatch)
	defer word.Release()
	oleutil.PutProperty(word, "DisplayAlerts", false)

	oleutil.PutProperty(word, "Visible", false)
	//oleutil.PutProperty(word, "Visible", true)

	documents := oleutil.MustGetProperty(word, "Documents").ToIDispatch()
	defer documents.Release()
	document := oleutil.MustCallMethod(documents, "Open", fileName).ToIDispatch()
	//document := oleutil.MustCallMethod(documents, "OpenNoRepairDialog", fileName).ToIDispatch()
	//document := oleutil.MustCallMethod(documents, "OpenNoRepairDialog", fileName, false, true).ToIDispatch()

	defer document.Release()

	oleutil.MustCallMethod(document, "SaveAs2", pdfPath, 17).ToIDispatch()
	oleutil.CallMethod(document, "Close")
	oleutil.CallMethod(word, "Quit")

	log.Info("officeWord2pdf - success")
}

func officeExcel2pdf(fileName string, pdfPath string) {
	log.Info("officeExcel2pdf - start")
	log.Info("fileName=" + fileName)
	log.Info("pdfPath=" + pdfPath)

	err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	if err != nil {
		log.Error(err)
		return
	}
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		log.Error(err)
		return
	}
	defer unknown.Release()

	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Error(err)
		return
	}
	defer excel.Release()

	oleutil.PutProperty(excel, "DisplayAlerts", false)
	oleutil.PutProperty(excel, "Visible", false)

	workbooks_ole, err := oleutil.GetProperty(excel, "Workbooks")
	if err != nil {
		log.Error(err)
		return
	}
	workbooks := workbooks_ole.ToIDispatch()
	defer workbooks.Release()

	workbook_ole, err := oleutil.CallMethod(workbooks, "Open", fileName, true)
	if err != nil {
		log.Error(err)
		return
	}
	workbook := workbook_ole.ToIDispatch()
	defer workbook.Release()

	worksheet_ole, err := oleutil.GetProperty(workbook, "Worksheets", 1)
	if err != nil {
		log.Error(err)
		return
	}
	worksheet := worksheet_ole.ToIDispatch()
	defer worksheet.Release()

	_, err = oleutil.CallMethod(worksheet, "ExportAsFixedFormat", 0, pdfPath)
	if err != nil {
		log.Error(err)
	}
	oleutil.PutProperty(workbook, "Saved", true)
	oleutil.CallMethod(workbook, "Close")
	oleutil.CallMethod(excel, "Quit")

	log.Info("officeExcel2pdf - success")
}

func officePpt2pdf(fileName string, pdfPath string) {
	log.Info("officePpt2pdf - start")

	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
	defer ole.CoUninitialize()
	unknown, err := oleutil.CreateObject("PowerPoint.Application")
	if err != nil {
		checkErr(err)
		return
	}
	defer unknown.Release()
	ppt, _ := unknown.QueryInterface(ole.IID_IDispatch)
	defer ppt.Release()
	presentations := oleutil.MustGetProperty(ppt, "Presentations").ToIDispatch()
	defer presentations.Release()
	presentation := oleutil.MustCallMethod(presentations, "Open", fileName).ToIDispatch()
	defer presentation.Release()
	oleutil.MustCallMethod(presentation, "SaveAs", pdfPath, 32).ToIDispatch()
	oleutil.CallMethod(presentation, "Close")
	oleutil.CallMethod(ppt, "Quit")

	log.Info("officePpt2pdf - success")
}
