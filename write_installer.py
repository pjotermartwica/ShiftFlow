"""Zapisuje install.ps1 bez problemow z kodowaniem."""
import os

ps1 = r"""# install.ps1 - Instalator ShiftFlow
# Uruchom: prawy klik na pliku -> "Uruchom przy uzyciu programu Windows PowerShell"
# Nie wymaga uprawnien Administratora.

$ErrorActionPreference = "Stop"

$AppName    = "ShiftFlow"
$InstallDir = "$env:LOCALAPPDATA\$AppName"
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ExeSrc     = Join-Path $ScriptDir "dist\ShiftFlow.exe"
$ExeDest    = "$InstallDir\ShiftFlow.exe"
$IcoSrc     = Join-Path $ScriptDir "icon.ico"
$IcoDest    = "$InstallDir\icon.ico"
$EnvSrc     = Join-Path $ScriptDir ".env"

Write-Host ""
Write-Host "=== Instalacja ShiftFlow ===" -ForegroundColor Cyan

# 1 - Kopiowanie plikow
Write-Host "[1/4] Kopiowanie plikow do $InstallDir ..." -ForegroundColor Yellow
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
Copy-Item $ExeSrc $ExeDest -Force
Copy-Item $IcoSrc $IcoDest -Force
if (Test-Path $EnvSrc) {
    Copy-Item $EnvSrc "$InstallDir\.env" -Force
    Write-Host "    .env skopiowany (klucz API)." -ForegroundColor DarkGreen
} else {
    Write-Host "    Brak pliku .env." -ForegroundColor DarkYellow
}
Write-Host "    OK" -ForegroundColor Green

# 2 - Skrot na pulpicie
Write-Host "[2/4] Tworzenie skrotu na pulpicie ..." -ForegroundColor Yellow
$DesktopPath  = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = "$DesktopPath\ShiftFlow.lnk"
$Shell    = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath       = $ExeDest
$Shortcut.WorkingDirectory = $InstallDir
$Shortcut.IconLocation     = "$IcoDest,0"
$Shortcut.Description      = "ShiftFlow - harmonogram pracy z AI"
$Shortcut.Save()
Write-Host "    OK - $ShortcutPath" -ForegroundColor Green

# 3 - Rejestracja rozszerzenia .grafik (HKCU - bez admina)
Write-Host "[3/4] Rejestracja rozszerzenia .grafik ..." -ForegroundColor Yellow
$Base    = "HKCU:\Software\Classes"
$RegExt  = "$Base\.grafik"
$RegType = "$Base\ShiftFlow.Document"

New-Item -Path $RegExt  -Force | Out-Null
Set-ItemProperty -Path $RegExt -Name "(default)"     -Value "ShiftFlow.Document"
Set-ItemProperty -Path $RegExt -Name "Content Type"  -Value "application/x-grafik"
Set-ItemProperty -Path $RegExt -Name "PerceivedType" -Value "Document"

New-Item -Path $RegType -Force | Out-Null
Set-ItemProperty -Path $RegType -Name "(default)" -Value "ShiftFlow"

New-Item -Path "$RegType\DefaultIcon" -Force | Out-Null
Set-ItemProperty -Path "$RegType\DefaultIcon" -Name "(default)" -Value "`"$IcoDest`",0"

New-Item -Path "$RegType\shell\open\command" -Force | Out-Null
Set-ItemProperty -Path "$RegType\shell\open\command" -Name "(default)" -Value "`"$ExeDest`" `"%1`""

Write-Host "    OK" -ForegroundColor Green

# 4 - Dodawanie do menu Nowy
Write-Host "[4/4] Dodawanie do menu 'Nowy' ..." -ForegroundColor Yellow
New-Item -Path "$RegType\ShellNew" -Force | Out-Null
Set-ItemProperty -Path "$RegType\ShellNew" -Name "NullFile" -Value ""
Set-ItemProperty -Path "$RegType\ShellNew" -Name "MenuText" -Value "ShiftFlow"
Set-ItemProperty -Path "$RegType\ShellNew" -Name "IconPath" -Value $IcoDest

# Odswiezenie Eksploratora Windows
$sig = '[System.Runtime.InteropServices.DllImport("shell32.dll")] public static extern void SHChangeNotify(int e, int f, IntPtr a, IntPtr b);'
Add-Type -MemberDefinition $sig -Name WinShell -Namespace Shell32 -ErrorAction SilentlyContinue
[Shell32.WinShell]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
Write-Host "    OK" -ForegroundColor Green

Write-Host ""
Write-Host "Instalacja zakonczona pomyslnie!" -ForegroundColor Cyan
Write-Host "  Lokalizacja : $ExeDest"
Write-Host "  Skrot       : $ShortcutPath"
Write-Host "  Rozszerzenie: .grafik"
Write-Host ""
Write-Host "Jesli pozycja 'Nowy' nie pojawia sie od razu - wyloguj sie i zaloguj ponownie." -ForegroundColor DarkYellow
Write-Host ""
Read-Host "Nacisnij Enter, aby zamknac"
"""

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "install.ps1")
with open(out, "w", encoding="utf-8-sig") as f:
    # utf-8-sig adds BOM which makes PowerShell treat the file as UTF-8
    # correctly on all Windows versions
    f.write(ps1)
print(f"Zapisano: {out}")
