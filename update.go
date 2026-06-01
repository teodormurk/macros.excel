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
		fatal("modules/ not found")
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
	fmt.Println("Opening workbook ...")
	exec.Command("osascript", "-e",
		fmt.Sprintf(`tell application "Microsoft Excel" to open POSIX file "%s"`, absXlsm),
	).Run()

	fmt.Println("Waiting 3s ...")
	exec.Command("osascript", "-e", "delay 3").Run()

	out, err := exec.Command("osascript", "-e",
		`tell application "Microsoft Excel" to run VBMacro "InstallModules"`).CombinedOutput()
	if err != nil {
		fmt.Println(string(out))
		fmt.Println()
		fmt.Println("Installer not found. Import Installer.bas once:")
		fmt.Println("  Alt+F11 -> right-click -> Import File -> modules/Installer.bas")
		return
	}
	fmt.Println("Done!")
}

func updateWin(absXlsm string) {
	vbs := fmt.Sprintf(`Set excel = CreateObject("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
Set wb = excel.Workbooks.Open("%s")
On Error Resume Next
excel.Run "InstallModules"
If Err.Number <> 0 Then
    MsgBox "Installer.bas not found in workbook." & vbCrLf & vbCrLf & _
           "Import it once: Alt+F11 -> right-click -> Import File -> modules/Installer.bas" & vbCrLf & _
           "Then save and run update again.", vbExclamation, "Update"
    wb.Close False
    excel.Quit
    WScript.Quit 1
End If
wb.Save
wb.Close
excel.Quit`, absXlsm)

	tmp := filepath.Join(os.TempDir(), "update_macros.vbs")
	os.WriteFile(tmp, []byte(vbs), 0644)
	defer os.Remove(tmp)

	out, err := exec.Command("cscript", "//nologo", tmp).CombinedOutput()
	fmt.Print(string(out))
	if err != nil {
		fatal("cscript error: %s", err)
	}
	fmt.Println("Done!")
}

func fatal(format string, args ...any) {
	fmt.Fprintf(os.Stderr, format+"\n", args...)
	os.Exit(1)
}
