package main

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
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

	absXlsm, _ := filepath.Abs(xlsmPath)
	absDir := filepath.Dir(absXlsm)
	modAbsDir := absDir + "\\modules"

	vbs := `Set excel = CreateObject("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile("` + absXlsm + `")
shortPath = f.ShortPath
Set wb = excel.Workbooks.Open(shortPath)

On Error Resume Next
Set vb = wb.VBProject
errNum = Err.Number
On Error GoTo 0

If errNum <> 0 Then
    MsgBox "VBA project access denied." & vbCrLf & vbCrLf & _
           "Excel -> Options -> Trust Center -> Macro Settings" & vbCrLf & _
           "Check [x] Trust access to the VBA project object model", vbExclamation, "Error"
    wb.Close False
    excel.Quit
    WScript.Quit 1
End If

On Error Resume Next
vb.VBComponents.Remove vb.VBComponents("ModSubstCore")
vb.VBComponents.Remove vb.VBComponents("ModSubstUtils")
vb.VBComponents.Remove vb.VBComponents("ModSubstUI")
On Error GoTo 0

On Error Resume Next
vb.VBComponents.Import "` + modAbsDir + `\ModSubstCore.bas"
vb.VBComponents.Import "` + modAbsDir + `\ModSubstUtils.bas"
vb.VBComponents.Import "` + modAbsDir + `\ModSubstUI.bas"
importErr = Err.Number
On Error GoTo 0

If importErr <> 0 Then
    MsgBox "Failed to import modules. Check modules/ files exist.", vbExclamation, "Error"
    wb.Close False
    excel.Quit
    WScript.Quit 1
End If

twPath = "` + modAbsDir + `\ThisWorkbook.cls"
If fso.FileExists(twPath) Then
    Set cm = vb.VBComponents("ThisWorkbook").CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    fnum = FreeFile
    Open twPath For Input As fnum
    started = False
    Do Until EOF(fnum)
        Line Input #fnum, line
        If Not started Then
            If Left(Trim(line), 10) = "Attribute " Then
                started = True
            End If
        End If
        If started Then
            lineNum = lineNum + 1
            cm.InsertLines lineNum, line
        End If
    Loop
    Close fnum
End If

wb.Save
wb.Close
excel.Quit
WScript.Quit 0`

	tmp := filepath.Join(os.TempDir(), "update_macros.vbs")
	vbs = strings.ReplaceAll(vbs, "\n", "\r\n")
	os.WriteFile(tmp, utf8ToWin1251([]byte(vbs)), 0644)
	defer os.Remove(tmp)

	out, err := exec.Command("cscript", "//nologo", tmp).CombinedOutput()
	fmt.Print(string(out))
	if err != nil {
		fatal("cscript failed: %s", err)
	}
}

func utf8ToWin1251(data []byte) []byte {
	result := make([]byte, 0, len(data))
	for i := 0; i < len(data); {
		b := data[i]
		switch {
		case b < 0x80:
			result = append(result, b)
			i++
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
				i++
			}
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
	fmt.Fprintf(os.Stderr, format+"\n", args...)
	os.Exit(1)
}
