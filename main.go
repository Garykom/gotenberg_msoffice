package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/gofiber/fiber/v2"
	"github.com/gofiber/fiber/v2/log"
)

var config Config

func consoleConvert() {
	args := os.Args[1:]
	if len(args) != 2 {
		fmt.Println("Parameter: [Office source filepath] [PDF filepath]")
		return
	}
	docArr := []string{".doc", ".docx"}
	pptArr := []string{".ppt", ".pptx"}
	excelArr := []string{".xls", ".xlsx"}

	var fileName = args[0]
	var pdfPath = args[1]

	if !isFileExists(fileName) {
		fmt.Println("Source file does not exist")
		return
	}
	ext := strings.ToLower(filepath.Ext(fileName))

	if contains(docArr, ext) {
		//killApp("winword.exe")
		officeWord2pdf(fileName, pdfPath)
	} else if contains(pptArr, ext) {
		//killApp("powerpnt.exe")
		officePpt2pdf(fileName, pdfPath)
	} else if contains(excelArr, ext) {
		//killApp("excel.exe")
		officeExcel2pdf(fileName, pdfPath)
	} else {
		fmt.Println("Format not supported")
	}
}

func startServer() {
	killAppStart()

	app := fiber.New()

	app.Get("/health", getHealth)
	app.Post("/forms/libreoffice/convert", postFormsLibreOfficeConvert)
	app.Get("/config", getConfig)
	app.Get("/log", getLog)
	app.Get("/clean", getClean)

	listenAddres := ":" + strconv.Itoa(config.Port)
	if err := app.Listen(listenAddres); err != nil {
		log.Fatalf("Error starting server: %v", err)
	}
}

func main() {
	logInit()
	configDefault()
	configLoad()
	logSetLevel(config.LogLevel)

	args := os.Args[1:]
	if len(args) == 0 {
		startServer()
	} else {
		consoleConvert()
	}
}
