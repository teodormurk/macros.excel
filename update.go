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
	dir, _ := os.Getwd()
	xlsmPath := filepath.Join(dir, xlsmName)
	modDir := filepath.Join(dir, modulesDir)

	if _, err := os.Stat(xlsmPath); os.IsNotExist(err) {
		fatal("%s not found", xlsmName)
	}
	if _, err := os.Stat(modDir); os.IsNotExist(err) {
		fatal("modules/ directory not found")
	}

	fmt.Printf("Updating %s ...\n", xlsmName)
	fmt.Println(strings.Repeat("=", 40))

	absXlsm, _ := filepath.Abs(xlsmPath)

	switch runtime.GOOS {
	case "darwin":
		updateMac(absXlsm)
	case "windows":
		updateWin(absXlsm)
	default:
		fatal("unsupported platform")
	}
}

func updateMac(absXlsm string) {
	fmt.Println("Opening workbook in Excel ...")
	out, err := exec.Command("osascript", "-e",
		fmt.Sprintf(`tell application "Microsoft Excel" to open POSIX file "%s"`, absXlsm),
	).CombinedOutput()
	if err != nil {
		fmt.Printf("open: %s: %s\n", err, out)
	}
	fmt.Println("Waiting 3s for Excel to load ...")
	out, _ = exec.Command("osascript", "-e", "delay 3").CombinedOutput()

	macros := []string{
		`tell application "Microsoft Excel" to run VBMacro "InstallModules"`,
		`tell application "Microsoft Excel" to do Visual Basic "InstallModules"`,
	}

	for _, cmd := range macros {
		fmt.Printf("Trying: %s\n", cmd)
		out, err := exec.Command("osascript", "-e", cmd).CombinedOutput()
		if err != nil {
			fmt.Printf("  -> error: %s\n", strings.TrimSpace(string(out)))
			continue
		}
		fmt.Println("  -> OK!")
		return
	}

	fmt.Println()
	fmt.Println("=== MANUAL INSTALLATION ===")
	fmt.Println("Installer.bas was not found in the workbook.")
	fmt.Println("First-time setup (only once):")
	fmt.Println("1. Open", xlsmName, "in Excel")
	fmt.Println("2. Press Alt+F11 (or Option+F11) to open VBA Editor")
	fmt.Println("3. Right-click on VBAProject -> Import File")
	fmt.Println("4. Select modules/Installer.bas")
	fmt.Println("5. Ctrl+S to save")
	fmt.Println("6. Close and reopen the workbook")
	fmt.Println("7. Run ./update again")
	fmt.Println()
	fmt.Println("Or import all modules manually:")
	fmt.Println("  modules/ModSubstCore.bas")
	fmt.Println("  modules/ModSubstUtils.bas")
	fmt.Println("  modules/ModSubstUI.bas")
	fmt.Println("  modules/ThisWorkbook.cls (paste code)")
}

func updateWin(absXlsm string) {
	var sb strings.Builder
	sb.WriteString(fmt.Sprintf(`Set wb = GetObject("%s")`, absXlsm))
	sb.WriteString("\nOn Error Resume Next\n")
	sb.WriteString("wb.Application.Run \"InstallModules\"\n")
	sb.WriteString("If Err.Number <> 0 Then\n")
	sb.WriteString("  WScript.Echo \"ERROR: Installer macro not found. Import Installer.bas first.\"\n")
	sb.WriteString("End If\n")
	tmp := filepath.Join(os.TempDir(), "run_installer.vbs")
	os.WriteFile(tmp, []byte(sb.String()), 0644)
	defer os.Remove(tmp)
	out, err := exec.Command("cscript", "//nologo", tmp).CombinedOutput()
	fmt.Print(string(out))
	if err != nil {
		fatal("cscript: %s", err)
	}
	fmt.Println("Done!")
}

func fatal(format string, args ...any) {
	fmt.Fprintf(os.Stderr, format+"\n", args...)
	os.Exit(1)
}
