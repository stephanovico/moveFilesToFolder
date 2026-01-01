# Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
# .\deploy.ps1

#Requires -RunAsAdministrator

param(
    [switch]$Uninstall
)

$InstallDir = "C:\deploy\MoveToNewFolder"
$ExeName = "moveFilesToFolder.exe"

function Test-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-Error "This script must be run as Administrator"
    exit 1
}

if ($Uninstall) {
    Write-Host "Uninstalling..." -ForegroundColor Yellow
    
    # Remove SendTo shortcut
    $SendToPath = [Environment]::GetFolderPath("SendTo")
    $ShortcutPath = "$SendToPath\Move to New Folder.lnk"
    if (Test-Path $ShortcutPath) {
        Remove-Item $ShortcutPath -Force
        Write-Host "Removed SendTo shortcut" -ForegroundColor Green
    }
    
    Write-Host "Uninstall complete. You can now delete: $InstallDir" -ForegroundColor Green
    exit 0
}

Write-Host "Building project..." -ForegroundColor Cyan
dotnet publish -c Release -r win-x64 --self-contained false

if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed"
    exit 1
}

$publishPath = "bin\Release\net10.0-windows\win-x64\publish"

if (-not (Test-Path $publishPath)) {
    Write-Error "Publish folder not found: $publishPath"
    exit 1
}

Write-Host "Creating install directory: $InstallDir" -ForegroundColor Cyan
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null

Write-Host "Copying files..." -ForegroundColor Cyan
Copy-Item "$publishPath\*" $InstallDir -Recurse -Force

$exePath = Join-Path $InstallDir $ExeName
if (-not (Test-Path $exePath)) {
    Write-Error "Executable not found after copy: $exePath"
    exit 1
}

Write-Host "Creating SendTo shortcut..." -ForegroundColor Cyan
$SendToPath = [Environment]::GetFolderPath("SendTo")
$ShortcutPath = "$SendToPath\Move to New Folder.lnk"

# Remove old shortcut if exists
if (Test-Path $ShortcutPath) {
    Remove-Item $ShortcutPath -Force
}

# Create new shortcut
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $exePath
$Shortcut.WorkingDirectory = $InstallDir
$Shortcut.Description = "Move selected files/folders to a new folder"
$Shortcut.IconLocation = "shell32.dll,3"
$Shortcut.Save()

# Release COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WshShell) | Out-Null

Write-Host ""
Write-Host "==================================" -ForegroundColor Green
Write-Host "Deployment complete!" -ForegroundColor Green
Write-Host "==================================" -ForegroundColor Green
Write-Host ""
Write-Host "Installed to: $InstallDir" -ForegroundColor Cyan
Write-Host "SendTo shortcut: $ShortcutPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "Usage:" -ForegroundColor Yellow
Write-Host "  1. Select one or more files/folders" -ForegroundColor White
Write-Host "  2. Right-click -> Send To -> Move to New Folder" -ForegroundColor White
Write-Host "  3. Enter folder name" -ForegroundColor White
Write-Host "  4. Done!" -ForegroundColor White
Write-Host ""
Write-Host "To uninstall: .\deploy.ps1 -Uninstall" -ForegroundColor Gray