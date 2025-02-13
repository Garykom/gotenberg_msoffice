package main

import (
	"encoding/json"
	"os"
	"strings"

	"github.com/gofiber/fiber/v2/log"
)

type Extensions struct {
	Word       []string `json:"ms_word"`
	Excel      []string `json:"ms_excel"`
	PowerPoint []string `json:"ms_powerpoint"`
	PDF        []string `json:"pdf"`
}

type Config struct {
	Comments       string     `json:"comments"`
	Port           int        `json:"port"`
	LogLevel       string     `json:"log_level"`
	KillAppTimeout int        `json:"kill_app_timeout"`
	Extensions     Extensions `json:"file_extensions"`
}

func configDefault() {
	config.Comments = ""
	config.Comments += "Available log_level: Trace, Debug, Info, Warn, Error\n"
	config.Comments += "KillAppTimeout: Timeout in sec to kill winword.exe (excel.exe, powerpoint.exe), 0 - disabled \n"
	config.Port = 3000
	config.LogLevel = "Info"
	config.KillAppTimeout = 0

	word := ".doc .docm .docx .dot .dotm .dotx .htm .html .mht .mhtml .odt .rtf .txt .wps .xml .xps"
	config.Extensions.Word = strings.Split(word, " ")

	excel := ".csv .dbf .ods .xls .xlsb .xlsm .xlsx .xlt .xltm .xltx .xlw"
	config.Extensions.Excel = strings.Split(excel, " ")

	powerpoint := ".odp .ppt .pptx"
	config.Extensions.PowerPoint = strings.Split(powerpoint, " ")

	config.Extensions.PDF = []string{".pdf"}
}

func configSave() {
	configFilePath := getConfigFilePath()
	file, _ := json.MarshalIndent(&config, "", " ")
	err := os.WriteFile(configFilePath, file, 0644)
	checkErr(err)
}

func configInit() {
	configDefault()
	configSave()
}

func configLoad() {
	configFilePath := getConfigFilePath()
	if !isFileExists(configFilePath) {
		configInit()
		return
	}
	bufferJSON, err := os.ReadFile(configFilePath)
	if err != nil {
		log.Error(err)
		configInit()
		return
	}
	err = json.Unmarshal(bufferJSON, &config)
	if err != nil {
		log.Error(err)
		configInit()
		return
	}
}
