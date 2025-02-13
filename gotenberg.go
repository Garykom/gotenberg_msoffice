package main

import (
	"archive/zip"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/gofiber/fiber/v2"
	"github.com/gofiber/fiber/v2/log"
)

func DownloadNoEscape(c *fiber.Ctx, file string, filename ...string) error {
	var fname string
	if len(filename) > 0 {
		fname = filename[0]
	} else {
		fname = filepath.Base(file)
	}
	c.Set(fiber.HeaderContentDisposition, `attachment; filename="`+fname+`"`)
	return c.SendFile(file)
}

func postFormsLibreOfficeConvert(c *fiber.Ctx) error {

	log.Info("post /forms/libreoffice/convert")

	tempDir, err := os.MkdirTemp("", "gotenberg_msoffice_")
	defer os.RemoveAll(tempDir)
	if err != nil {
		log.Error("Error in uploading File : ", err)
	}
	log.Debug("Temp dir name: ", tempDir)

	form, err := c.MultipartForm()
	if err != nil {
		return err
	}

	// Файлы
	files := form.File["files"]
	var filePaths []string
	for fileIndex, fileHeader := range files {
		fileName := fileHeader.Filename
		log.Debug(fileName, fileHeader.Size, fileHeader.Header["Content-Type"][0])

		extension := filepath.Ext(fileName)
		log.Debug("extension = " + extension)

		tempFileTemplate := "file_" + strconv.Itoa(fileIndex) + "_*" + extension
		tempFilePath, _ := getTempFilePath(tempDir, tempFileTemplate)
		log.Debug("tempFilePath = " + tempFilePath)

		filePaths = append(filePaths, tempFilePath)

		err = c.SaveFile(fileHeader, tempFilePath)
		defer os.Remove(tempFilePath)
		if err != nil {
			return err
		}
	}

	// Параметры
	merge := false
	mergeValues := form.Value["merge"]
	for _, value := range mergeValues {
		if value == "true" {
			merge = true
		} else if value == "false" {
			merge = false
		} else {
			log.Error("Непонятное значение merge = \"" + value + "\"! Используем merge = false")
			merge = false
		}
	}
	log.Debug("merge = " + strconv.FormatBool(merge))

	// Конвертация
	var pdfFilePaths []string

	for fileIndex, filePath := range filePaths {
		log.Debug("fileIndex = " + strconv.Itoa(fileIndex))

		fileDir := filepath.Dir(filePath)
		log.Debug(fileDir)
		fileName := filepath.Base(filePath)
		log.Debug(fileName)
		fileExt := filepath.Ext(filePath)
		log.Debug(fileExt)

		fileNameWithoutExtension := strings.TrimSuffix(fileName, fileExt)

		pdfFileName := fileNameWithoutExtension + ".pdf"
		pdfFilePath := filepath.Join(fileDir, pdfFileName)
		log.Debug(pdfFilePath)

		extWord := config.Extensions.Word
		extExcel := config.Extensions.Excel
		extPowerPoint := config.Extensions.PowerPoint
		extPDF := config.Extensions.PDF

		ext := strings.ToLower(fileExt)

		if contains(extWord, ext) {
			officeWord2pdf(filePath, pdfFilePath)
		} else if contains(extPowerPoint, ext) {
			officePpt2pdf(filePath, pdfFilePath)
		} else if contains(extExcel, ext) {
			officeExcel2pdf(filePath, pdfFilePath)
		} else if contains(extPDF, ext) {
			// Ничего не делаем
		} else {
			log.Error("Format not supported: \"" + ext + "\"")
			continue
		}

		pdfFilePaths = append(pdfFilePaths, pdfFilePath)
	}

	resultFilePath := ""
	resultFileName := ""

	if merge {

		// Объединение в один PDF файл
		resultFileName = "result.pdf"
		resultFilePath = filepath.Join(tempDir, resultFileName)
		defer os.Remove(resultFilePath)

		appDir := getApplicationDirectory()
		qpdfDir := "\\qpdf-msvc32\\"
		qpdfName := "qpdf.exe"
		qpdfExePath := filepath.Join(appDir, qpdfDir, qpdfName)
		log.Debug("qpdfExePath = " + qpdfExePath)

		//qpdfArg := "--empty --pages --file=a.pdf --file=b.pdf --file=c.pdf -- out.pdf"
		var qpdfArg []string

		qpdfArg = append(qpdfArg, "--empty")
		qpdfArg = append(qpdfArg, "--pages")
		for _, pdfFilePath := range pdfFilePaths {
			qpdfArg = append(qpdfArg, "--file="+pdfFilePath)
		}
		qpdfArg = append(qpdfArg, "--")
		qpdfArg = append(qpdfArg, resultFilePath)

		log.Debug("Run QPDF")
		cmd := exec.Command(qpdfExePath, qpdfArg...)
		err = cmd.Run()
		checkErr(err)
		log.Debug("Complete QPDF")

	} else {

		// Сохранение всех PDF в один архив ZIP
		resultFileName = "result.zip"
		resultFilePath = filepath.Join(tempDir, resultFileName)
		defer os.Remove(resultFilePath)

		log.Debug("Make ZIP")

		zipFile, err := os.Create(resultFilePath)
		checkErr(err)
		zipWriter := zip.NewWriter(zipFile)
		for _, pdfFilePath := range pdfFilePaths {
			fileToZip, err := os.Open(pdfFilePath)
			checkErr(err)

			fileInfo, err := fileToZip.Stat()
			checkErr(err)
			header, err := zip.FileInfoHeader(fileInfo)
			checkErr(err)

			header.Name = filepath.Base(pdfFilePath)

			writer, err := zipWriter.CreateHeader(header)
			checkErr(err)

			_, err = io.Copy(writer, fileToZip)
			checkErr(err)

			fileToZip.Close()
		}
		zipWriter.Close()
		zipFile.Close()
		log.Debug("Complete ZIP")
	}

	log.Debug("resultFilePath = " + resultFilePath)

	resultFile, err := os.ReadFile(resultFilePath)
	checkErr(err)

	err = c.Send(resultFile)
	checkErr(err)

	err = os.Remove(resultFilePath)
	checkErr(err)

	return err
}

type JSONTimeNs struct {
	time.Time
}

func (ns JSONTimeNs) MarshalJSON() ([]byte, error) {
	str := fmt.Sprintf(`"%s"`, ns.Format("2006-01-02T15:04:05.000000000Z"))
	return []byte(str), nil
}

func (ns *JSONTimeNs) UnmarshalJSON(text []byte) error {
	if len(text) < 32 {
		return errors.New("malformed nanosecond timestamp")
	}
	str := string(text[1:31])
	t, err := time.Parse("2006-01-02T15:04:05.000000000Z", str)
	if err != nil {
		return err
	}
	ns.Time = t
	return nil
}

type HealthModuleStatus struct {
	Status    string     `json:"status"`
	Timestamp JSONTimeNs `json:"timestamp"`
}

type HealthDetails struct {
	Chromium    HealthModuleStatus `json:"chromium"`
	LibreOffice HealthModuleStatus `json:"libreoffice"`
}

type Health struct {
	Status  string        `json:"status"`
	Details HealthDetails `json:"details"`
}

func getHealth(c *fiber.Ctx) error {
	//{"status":"up","details":{"chromium":{"status":"up","timestamp":"2025-02-02T18:37:27.151303956Z"},"libreoffice":{"status":"up","timestamp":"2025-02-02T18:37:27.018348708Z"}}}

	log.Info("get /health")

	var health Health
	health.Status = "up"

	health.Details.Chromium.Status = "down"
	health.Details.Chromium.Timestamp = JSONTimeNs{time.Now()}

	health.Details.LibreOffice.Status = "down"
	health.Details.LibreOffice.Timestamp = JSONTimeNs{time.Now()}
	_, err := json.Marshal(health)
	if err != nil {
		log.Error(err)
		return c.JSON(fiber.Map{"status": 200, "message": "Error", "data": err})
	}
	return c.JSON(health)
}

func getConfig(c *fiber.Ctx) error {
	log.Info("get /config")
	configFilePath := getConfigFilePath()
	buf, _ := os.ReadFile(configFilePath)
	_, err := c.WriteString(string(buf))
	return err
}

func getLog(c *fiber.Ctx) error {
	log.Info("get /log")
	logFilePath := getLogFilePath()
	buf, _ := os.ReadFile(logFilePath)
	_, err := c.WriteString(string(buf))
	return err
}

func getClean(c *fiber.Ctx) error {
	log.Info("get /clean")

	killApp("winword.exe")
	killApp("excel.exe")
	killApp("powerpnt.exe")

	resultString := "Processes killed"
	_, err := c.WriteString(resultString)
	return err
}
