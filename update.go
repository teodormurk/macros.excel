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
	xlsmName   = "\u041a\u043d\u0438\u0433\u04301.xlsm"
	modulesDir = "modules"
)

func main() {
	dir, err := os.Getwd()
	if err != nil {
		fatal("\u043d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043e\u043f\u0440\u0435\u0434\u0435\u043b\u0438\u0442\u044c \u0442\u0435\u043a\u0443\u0449\u0443\u044e \u0434\u0438\u0440\u0435\u043a\u0442\u043e\u0440\u0438\u044e: %v", err)
	}

	xlsmPath := filepath.Join(dir, xlsmName)
	modDir := filepath.Join(dir, modulesDir)

	if _, err := os.Stat(xlsmPath); os.IsNotExist(err) {
		fatal("%s \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d \u0440\u044f\u0434\u043e\u043c \u0441\u043e \u0441\u043a\u0440\u0438\u043f\u0442\u043e\u043c", xlsmName)
	}
	if _, err := os.Stat(modDir); os.IsNotExist(err) {
		fatal("\u043f\u0430\u043f\u043a\u0430 modules/ \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u0430")
	}

	fmt.Printf("Updating modules in %s\n", xlsmName)
	fmt.Println(strings.Repeat("=", 40))

	absXlsm, _ := filepath.Abs(xlsmPath)

	switch runtime.GOOS {
	case "darwin":
		runAS(fmt.Sprintf(`tell application "Microsoft Excel" to open POSIX file "%s"`, absXlsm))
		// Wait for Excel to load the workbook
		runAS(`delay 2`)
		out, err := exec.Command("osascript", "-e", `tell application "Microsoft Excel" to do Visual Basic "InstallModules"`).CombinedOutput()
		if err != nil {
			fatal("Excel error: %s\n%s", err, string(out))
		}
		fmt.Println("Done!")

	case "windows":
		var sb strings.Builder
		sb.WriteString(fmt.Sprintf(`Set wb = GetObject("%s")`, absXlsm))
		sb.WriteString("\nwb.Application.Run \"InstallModules\"\n")
		tmp := filepath.Join(os.TempDir(), "run_installer.vbs")
		os.WriteFile(tmp, []byte(sb.String()), 0644)
		defer os.Remove(tmp)
		out, err := exec.Command("cscript", "//nologo", tmp).CombinedOutput()
		if err != nil {
			fatal("Excel error: %s\n%s", err, string(out))
		}
		fmt.Println("Done!")

	default:
		fatal("platform %s is not supported", runtime.GOOS)
	}
}

func runAS(script string) {
	out, err := exec.Command("osascript", "-e", script).CombinedOutput()
	if err != nil {
		fmt.Printf("AppleScript: %s\n%s\n", err, string(out))
	}
}

func fatal(format string, args ...any) {
	fmt.Fprintf(os.Stderr, "Error: "+format+"\n", args...)
	os.Exit(1)
}
