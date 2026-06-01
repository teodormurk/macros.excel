package main

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
)

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
		abs, _ := filepath.Abs(p)
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
		abs, _ := filepath.Abs(p)
		sb.WriteString(fmt.Sprintf("On Error Resume Next\n"))
		sb.WriteString(fmt.Sprintf("For Each c In vb.VBComponents\n"))
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

func fatal(format string, args ...any) {
	fmt.Fprintf(os.Stderr, "Ошибка: "+format+"\n", args...)
	os.Exit(1)
}
