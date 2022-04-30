#Requires -RunAsAdministrator

param (
    [Parameter(Mandatory = $true)]
    [string]
    $Version
)

# create /bin if not existing
if (-not(Test-Path ".\bin")) {
    New-Item -Name "bin" -ItemType Directory
}

# clear the dir
Remove-Item -Path ".\bin\*"

# download tool if necessary
if (-not(Get-Module ps2exe -ListAvailable)) {
    Install-Module ps2exe
}

$year = Get-Date -Format "yyyy"

# build executable
Invoke-PS2EXE -inputFile ".\MMCSetupTool.ps1" -outputFile ".\bin\MMCSetupTool.exe" -iconFile ".\assets\mmc.ico" -title "MMC Setup Tool" -company "Multimedia Center Veenendaal" -copyright "$year MMC Store BV" -version "$Version" -requireAdmin 

# zip the files
Compress-Archive -Path ".\bin\MMCSetupTool.exe", ".\assets" -CompressionLevel Fastest -DestinationPath ".\bin\MMCSetupTool.zip"
