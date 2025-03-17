package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/gofiber/fiber/v2/log"
)

func officeWord2pdf(fileName string, pdfPath string) error {
	log.Info("officeWord2pdf - start")
	log.Info("fileName=" + fileName)
	log.Info("pdfPath=" + pdfPath)

	ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("Word.Application")
	if err != nil {
		log.Error(err)
		return err
	}
	defer unknown.Release()

	word, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Error(err)
		return err
	}
	defer word.Release()

	oleutil.PutProperty(word, "DisplayAlerts", false)
	oleutil.PutProperty(word, "Visible", false)

	documents_ole, err := oleutil.GetProperty(word, "Documents")
	if err != nil {
		log.Error(err)
		return err
	}
	documents := documents_ole.ToIDispatch()
	defer documents.Release()

	document_ole, err := oleutil.CallMethod(documents, "Open", fileName)
	if err != nil {
		log.Error(err)
		return err
	}
	document := document_ole.ToIDispatch()
	defer document.Release()

	_, err = oleutil.CallMethod(document, "SaveAs2", pdfPath, 17)
	if err != nil {
		log.Error(err)
	}
	oleutil.CallMethod(document, "Close")
	oleutil.CallMethod(word, "Quit")

	log.Info("officeWord2pdf - success")
	return err
}

func officeExcel2pdf(fileName string, pdfPath string) error {
	log.Info("officeExcel2pdf - start")
	log.Info("fileName=" + fileName)
	log.Info("pdfPath=" + pdfPath)

	err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	if err != nil {
		log.Error(err)
		return err
	}
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		log.Error(err)
		return err
	}
	defer unknown.Release()

	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Error(err)
		return err
	}
	defer excel.Release()

	oleutil.PutProperty(excel, "DisplayAlerts", false)
	oleutil.PutProperty(excel, "Visible", false)

	workbooks_ole, err := oleutil.GetProperty(excel, "Workbooks")
	if err != nil {
		log.Error(err)
		return err
	}
	workbooks := workbooks_ole.ToIDispatch()
	defer workbooks.Release()

	workbook_ole, err := oleutil.CallMethod(workbooks, "Open", fileName, true)
	if err != nil {
		log.Error(err)
		return err
	}
	workbook := workbook_ole.ToIDispatch()
	defer workbook.Release()

	worksheet_ole, err := oleutil.GetProperty(workbook, "Worksheets", 1)
	if err != nil {
		log.Error(err)
		return err
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
	return err
}

func officePpt2pdf(fileName string, pdfPath string) error {
	log.Info("officePpt2pdf - start")
	log.Info("fileName=" + fileName)
	log.Info("pdfPath=" + pdfPath)

	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("PowerPoint.Application")
	if err != nil {
		log.Error(err)
		return err
	}
	defer unknown.Release()

	ppt, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Error(err)
		return err
	}
	defer ppt.Release()

	presentations_ole, err := oleutil.GetProperty(ppt, "Presentations")
	if err != nil {
		log.Error(err)
		return err
	}
	presentations := presentations_ole.ToIDispatch()
	defer presentations.Release()

	presentation_ole, err := oleutil.CallMethod(presentations, "Open", fileName)
	if err != nil {
		log.Error(err)
		return err
	}
	presentation := presentation_ole.ToIDispatch()
	defer presentation.Release()

	_, err = oleutil.CallMethod(presentation, "SaveAs", pdfPath, 32)
	if err != nil {
		log.Error(err)
	}
	oleutil.CallMethod(presentation, "Close")
	oleutil.CallMethod(ppt, "Quit")

	log.Info("officePpt2pdf - success")
	return err
}
