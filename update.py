#!/usr/bin/env python3
"""
update.py — заменяет VBA-модули в Книга1.xlsm на версии из modules/.

Использование:
    python3 update.py
"""

import os
import sys
import platform

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MODULES_DIR = os.path.join(SCRIPT_DIR, "modules")
XLSM_NAME = "Книга1.xlsm"
XLSM_PATH = os.path.join(SCRIPT_DIR, XLSM_NAME)

BAS_MODULES = {
    "ModSubstCore.bas": "ModSubstCore",
    "ModSubstUtils.bas": "ModSubstUtils",
    "ModSubstUI.bas": "ModSubstUI",
}
THISWORKBOOK_FILE = "ThisWorkbook.cls"


def update_windows():
    try:
        import win32com.client
    except ImportError:
        print("Установка pywin32...")
        os.system(f"{sys.executable} -m pip install pywin32 --quiet")
        import win32com.client

    if not os.path.isfile(XLSM_PATH):
        print(f"Ошибка: {XLSM_NAME} не найден рядом со скриптом")
        sys.exit(1)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(os.path.abspath(XLSM_PATH))
    vb = wb.VBProject

    for filename, mod_name in BAS_MODULES.items():
        filepath = os.path.join(MODULES_DIR, filename)
        if not os.path.isfile(filepath):
            print(f"Пропуск: {filename} не найден")
            continue
        try:
            for comp in list(vb.VBComponents):
                if comp.Name == mod_name:
                    vb.VBComponents.Remove(comp)
                    break
        except Exception:
            pass
        vb.VBComponents.Import(os.path.abspath(filepath))
        print(f"Обновлён: {mod_name}")

    tw_path = os.path.join(MODULES_DIR, THISWORKBOOK_FILE)
    if os.path.isfile(tw_path):
        for comp in vb.VBComponents:
            if comp.Name == "ThisWorkbook":
                cm = comp.CodeModule
                cm.DeleteLines(1, cm.CountOfLines)
                with open(tw_path, "r", encoding="utf-8") as f:
                    lines = f.readlines()
                code_lines = []
                started = False
                for line in lines:
                    s = line.strip()
                    if not started:
                        if s.startswith("Option "):
                            started = True
                        else:
                            continue
                    code_lines.append(line.rstrip("\n"))
                for i, cl in enumerate(code_lines, 1):
                    cm.InsertLines(i, cl)
                print("Обновлён: ThisWorkbook")
                break

    wb.Save()
    wb.Close()
    excel.Quit()
    print("Готово!")


def run_vba_line(line):
    cmd = f'tell application "Microsoft Excel"\\n    do Visual Basic "{line}"\\nend tell'
    os.system(f"osascript -e '{cmd}'")


def update_macos():
    if not os.path.isfile(XLSM_PATH):
        print(f"Ошибка: {XLSM_NAME} не найден рядом со скриптом")
        sys.exit(1)

    abs_xlsm = os.path.abspath(XLSM_PATH)
    abs_modules = os.path.abspath(MODULES_DIR).replace(" ", "\\ ")

    print(f"Открытие {XLSM_NAME}...")
    run_vba_line("")
    os.system(f"""osascript -e 'tell application "Microsoft Excel" to open POSIX file "{abs_xlsm}"'""")

    for filename, mod_name in BAS_MODULES.items():
        filepath = os.path.join(MODULES_DIR, filename)
        if not os.path.isfile(filepath):
            print(f"Пропуск: {filename} не найден")
            continue
        abs_f = os.path.abspath(filepath)
        run_vba_line(f"On Error Resume Next")
        run_vba_line(f"ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(\"{mod_name}\")")
        run_vba_line(f"On Error GoTo 0")
        run_vba_line(f"ThisWorkbook.VBProject.VBComponents.Import \"{abs_f}\"")
        print(f"Обновлён: {mod_name}")

    tw_path = os.path.join(MODULES_DIR, THISWORKBOOK_FILE)
    if os.path.isfile(tw_path):
        abs_tw = os.path.abspath(tw_path)
        with open(tw_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        code_lines = []
        started = False
        for line in lines:
            s = line.strip()
            if not started:
                if s.startswith("Option "):
                    started = True
                else:
                    continue
            code_lines.append(line.rstrip("\n"))
        cm_ref = "ThisWorkbook.VBProject.VBComponents(\"ThisWorkbook\").CodeModule"
        run_vba_line(f"{cm_ref}.DeleteLines 1, {cm_ref}.CountOfLines")
        for i, cl in enumerate(code_lines, 1):
            escaped = cl.replace('"', '""')
            run_vba_line(f'{cm_ref}.InsertLine {i}, "{escaped}"')
        print("Обновлён: ThisWorkbook")

    run_vba_line("ThisWorkbook.Save")
    print("Готово! Можно закрыть Excel.")


def main():
    print(f"Обновление модулей в {XLSM_NAME}")
    print("=" * 40)

    if platform.system() == "Windows":
        update_windows()
    elif platform.system() == "Darwin":
        update_macos()
    else:
        print(f"Платформа {platform.system()} не поддерживается")
        sys.exit(1)


if __name__ == "__main__":
    main()
