package main

import (
	"fmt"
	"io"
	"os"
	"os/user"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/gofiber/fiber/v2/log"
	"github.com/shirou/gopsutil/process"
)

func getApplicationDirectory() string {
	exePath, err := os.Executable()
	checkErr(err)
	exeDir := filepath.Dir(exePath)
	return exeDir
}

func getConfigFilePath() string {
	appDir := getApplicationDirectory()

	configFileName := "config.json"
	configFilePath := filepath.Join(appDir, configFileName)

	return configFilePath
}

func getLogFilePath() string {
	appPath, err := os.Executable()
	checkErr(err)

	appDir := filepath.Dir(appPath)
	appName := filepath.Base(appPath)
	appExt := filepath.Ext(appPath)
	appNameWoExt := strings.TrimSuffix(appName, appExt)

	logFileName := appNameWoExt + ".log"
	logFilePath := filepath.Join(appDir, logFileName)

	return logFilePath
}

func logInit() {
	logFilePath := getLogFilePath()
	file, _ := os.OpenFile(logFilePath, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	iw := io.MultiWriter(os.Stdout, file)
	log.SetOutput(iw)
	log.SetLevel(log.LevelError)
}

func logSetLevel(configLevel string) {
	// Trace, Debug, Info, Warn, Error
	switch configLevel {
	case "Trace":
		log.SetLevel(log.LevelTrace)
	case "Debug":
		log.SetLevel(log.LevelDebug)
	case "Info":
		log.SetLevel(log.LevelInfo)
	case "Warn":
		log.SetLevel(log.LevelWarn)
	case "Error":
		log.SetLevel(log.LevelError)
	default:
		log.SetLevel(log.LevelInfo)
	}
}

func contains(s []string, e string) bool {
	for _, v := range s {
		if v == e {
			return true
		}
	}
	return false
}

func isFileExists(filename string) bool {
	if _, err := os.Open(filename); err != nil {
		if os.IsNotExist(err) {
			return false
		}
	}
	return true
}

func killAppStart() {
	if config.KillAppTimeout < 1 {
		return
	}
	time.AfterFunc(10*time.Second, func() {
		log.Trace("killApp timer")
		killApp("winword.exe")
		killApp("excel.exe")
		killApp("powerpnt.exe")
	})
}

func killApp(appName string) {
	log.Debug("killApp")

	if config.KillAppTimeout < 1 {
		return
	}

	currentUser, err := user.Current()
	if err != nil {
		log.Error(err)
	}
	username := currentUser.Username

	fmt.Printf("Username is: %s\n", username)
	ps, err := process.Processes()
	if err != nil {
		log.Error(err)
		return
	}
	for _, v := range ps {
		vUsername, err := v.Username()
		if err != nil {
			log.Trace(err)
			continue
		}
		log.Debug(vUsername)
		if vUsername != username {
			continue
		}

		vName, err := v.Name()
		if err != nil {
			log.Error(err)
			continue
		}
		log.Debug(vName)

		if strings.EqualFold(vName, appName) {
			log.Info("Process Name=" + vName)

			vPid := int(v.Pid)
			log.Info("Process Pid=" + strconv.Itoa(vPid))
			p, err := os.FindProcess(vPid)
			if err != nil {
				log.Error(err)
				continue
			}
			vCreateTime, err := v.CreateTime()
			if err != nil {
				log.Error(err)
				continue
			}
			timeCreate := time.UnixMilli(vCreateTime)
			log.Info("CreateTime = " + timeCreate.String())
			duraction := time.Since(timeCreate)
			duractionSeconds := duraction.Seconds()
			if duractionSeconds > float64(config.KillAppTimeout) {
				log.Info("killing process %d ", vPid)
				err = p.Kill()
				if err != nil {
					log.Info("process: %s killed failed\n", vName)
					log.Error(err)
				} else {
					log.Info("process: %s killed success\n", vName)
				}
				sleep(1000)
			}
		}
	}
}

func sleep(millisecond int) {
	time.Sleep(time.Duration(millisecond) * time.Millisecond)
}

func checkErr(err error) {
	if err != nil {
		log.Error(err)
	}
}

func getTempFilePath(tempDir, tempFileTemplate string) (string, error) {
	tempFile, err := os.CreateTemp(tempDir, tempFileTemplate)
	if err != nil {
		return "", err
	}
	tempFile.Close()
	tempFilePath := tempFile.Name()
	return tempFilePath, nil
}
