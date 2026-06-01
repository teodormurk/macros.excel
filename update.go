package main

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
)

func init() {
	if runtime.GOOS == "windows" {
		exec.Command("chcp", "65001").Run()
	}
}

const (
	xlsmName   = "Книга1.xlsm"
	modulesDir = "modules"
	twbFile    = "ThisWorkbook.cls"
)

var basModules = map[string]string{
	"ModSubstCore.bas":  "ModSubstCore",
	"ModSubstUtils.bas": "ModSubstUtils",
	"ModSubstUI.bas":    "ModSubstUI",
}

func main() {
	dir, err := os.Getwd()
	if err != nil {
		fatal("не удалось определить текущую директорию: %v", err)
	}

	xlsmPath := filepath.Join(dir, xlsmName)
	modDir := filepath.Join(dir, modulesDir)

	if _, err := os.Stat(xlsmPath); os.IsNotExist(err) {
		fatal("%s не найден рядом со скриптом", xlsmName)
	}

	fmt.Printf("Обновление модулей в %s\n", xlsmName)
	fmt.Println(strings.Repeat("=", 40))

	switch runtime.GOOS {
	case "darwin":
		updateMac(xlsmPath, modDir)
	case "windows":
		updateWin(xlsmPath, modDir)
	default:
		fatal("платформа %s не поддерживается", runtime.GOOS)
	}
}

func updateMac(xlsmPath, modDir string) {
	absXlsm, _ := filepath.Abs(xlsmPath)

	fmt.Printf("Открытие %s...\n", xlsmName)
	as(fmt.Sprintf(`tell application "Microsoft Excel" to open POSIX file "%s"`, absXlsm))

	for file, name := range basModules {
		p := filepath.Join(modDir, file)
		if _, err := os.Stat(p); os.IsNotExist(err) {
			fmt.Printf("Пропуск: %s не найден\n", file)
			continue
		}
		tmp, err := convertToAnsi(p)
		if err != nil {
			fmt.Printf("Ошибка конвертации %s: %v\n", file, err)
			continue
		}
		defer os.Remove(tmp)
		abs, _ := filepath.Abs(tmp)
		vba("On Error Resume Next")
		vba(fmt.Sprintf(`ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("%s")`, name))
		vba("On Error GoTo 0")
		vba(fmt.Sprintf(`ThisWorkbook.VBProject.VBComponents.Import "%s"`, abs))
		fmt.Printf("Обновлён: %s\n", name)
	}

	twPath := filepath.Join(modDir, twbFile)
	if _, err := os.Stat(twPath); err == nil {
		lines := readCodeLines(twPath)
		cm := `ThisWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule`
		vba(fmt.Sprintf("%s.DeleteLines 1, %s.CountOfLines", cm, cm))
		for i, line := range lines {
			escaped := strings.ReplaceAll(line, `"`, `""`)
			vba(fmt.Sprintf(`%s.InsertLine %d, "%s"`, cm, i+1, escaped))
		}
		fmt.Println("Обновлён: ThisWorkbook")
	}

	vba("ThisWorkbook.Save")
	fmt.Println("Готово! Можно закрыть Excel.")
}

func updateWin(xlsmPath, modDir string) {
	var sb strings.Builder

	absXlsm, _ := filepath.Abs(xlsmPath)
	fmt.Printf("Открытие %s...\n", xlsmName)

	sb.WriteString("Set excel = CreateObject(\"Excel.Application\")\n")
	sb.WriteString("excel.Visible = False\n")
	sb.WriteString("excel.DisplayAlerts = False\n")
	sb.WriteString(fmt.Sprintf("Set wb = excel.Workbooks.Open(\"%s\")\n", absXlsm))
	sb.WriteString("Set vb = wb.VBProject\n\n")

	for file, name := range basModules {
		p := filepath.Join(modDir, file)
		if _, err := os.Stat(p); os.IsNotExist(err) {
			fmt.Printf("Пропуск: %s не найден\n", file)
			continue
		}
		tmp, err := convertToAnsi(p)
		if err != nil {
			fmt.Printf("Ошибка конвертации %s: %v\n", file, err)
			continue
		}
		defer os.Remove(tmp)
		abs, _ := filepath.Abs(tmp)
		sb.WriteString("On Error Resume Next\n")
		sb.WriteString("For Each c In vb.VBComponents\n")
		sb.WriteString(fmt.Sprintf("    If c.Name = \"%s\" Then vb.VBComponents.Remove c\n", name))
		sb.WriteString("Next\n")
		sb.WriteString("On Error GoTo 0\n")
		sb.WriteString(fmt.Sprintf("vb.VBComponents.Import \"%s\"\n\n", abs))
		fmt.Printf("Обновлён: %s\n", name)
	}

	twPath := filepath.Join(modDir, twbFile)
	if _, err := os.Stat(twPath); err == nil {
		lines := readCodeLines(twPath)
		sb.WriteString("Set cm = vb.VBComponents(\"ThisWorkbook\").CodeModule\n")
		sb.WriteString("cm.DeleteLines 1, cm.CountOfLines\n")
		for i, line := range lines {
			escaped := strings.ReplaceAll(line, `"`, `""`)
			sb.WriteString(fmt.Sprintf("cm.InsertLines %d, \"%s\"\n", i+1, escaped))
		}
		fmt.Println("Обновлён: ThisWorkbook")
	}

	sb.WriteString("\nwb.Save\n")
	sb.WriteString("wb.Close\n")
	sb.WriteString("excel.Quit\n")

	tmp := filepath.Join(os.TempDir(), "update_macros.vbs")
	os.WriteFile(tmp, []byte(sb.String()), 0644)
	defer os.Remove(tmp)

	exec.Command("cscript", "//nologo", tmp).Run()
	fmt.Println("Готово!")
}

func vba(line string) {
	script := fmt.Sprintf(`tell application "Microsoft Excel" to do Visual Basic "%s"`, line)
	exec.Command("osascript", "-e", script).Run()
}

func as(script string) {
	exec.Command("osascript", "-e", script).Run()
}

func readCodeLines(path string) []string {
	data, _ := os.ReadFile(path)
	var lines []string
	started := false
	for _, line := range strings.Split(string(data), "\n") {
		t := strings.TrimSpace(line)
		if !started {
			if strings.HasPrefix(t, "Option ") {
				started = true
			} else {
				continue
			}
		}
		lines = append(lines, strings.TrimRight(line, "\r"))
	}
	return lines
}

func convertToAnsi(srcPath string) (string, error) {
	data, err := os.ReadFile(srcPath)
	if err != nil {
		return "", err
	}
	converted := utf8ToWin1251(data)
	tmp := filepath.Join(os.TempDir(), filepath.Base(srcPath)+".ansi")
	err = os.WriteFile(tmp, converted, 0644)
	return tmp, err
}

func utf8ToWin1251(data []byte) []byte {
	result := make([]byte, 0, len(data))
	for i := 0; i < len(data); {
		b := data[i]
		switch {
		case b < 0x80:
			result = append(result, b)
			i++
		case b == 0xC2:
			if i+1 < len(data) {
				b2 := data[i+1]
				i += 2
				switch b2 {
				case 0x85:
					result = append(result, 0xA5)
				case 0x89:
					result = append(result, 0xA9)
				case 0x8A:
					result = append(result, 0xAA)
				default:
					result = append(result, '?')
				}
			} else {
				result = append(result, '?')
				i++
			}
		case b == 0xD0:
			if i+1 < len(data) {
				b2 := data[i+1]
				i += 2
				switch {
				case b2 == 0x81:
					result = append(result, 0xA1)
				case b2 >= 0x90 && b2 <= 0xBF:
					result = append(result, b2+0x30)
				default:
					result = append(result, '?')
				}
			} else {
				result = append(result, '?')
				i++
			}
		case b == 0xD1:
			if i+1 < len(data) {
				b2 := data[i+1]
				i += 2
				switch {
				case b2 == 0x91:
					result = append(result, 0xB8)
				case b2 >= 0x80 && b2 <= 0x8F:
					result = append(result, b2+0x70)
				default:
					result = append(result, '?')
				}
			} else {
				result = append(result, '?')
				i++
			}
		case b == 0xE2:
			if i+2 < len(data) && data[i+1] == 0x80 && data[i+2] == 0x99 {
				result = append(result, 0x99)
				i += 3
			} else {
				result = append(result, '?')
				i++
			}
		default:
			result = append(result, '?')
			i++
		}
	}
	return result
}

func fatal(format string, args ...any) {
	fmt.Fprintf(os.Stderr, "Ошибка: "+format+"\n", args...)
	os.Exit(1)
}
