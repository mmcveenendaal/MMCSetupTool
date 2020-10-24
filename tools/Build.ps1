#Requires -RunAsAdministrator

param (
    [Parameter(Mandatory = $true)]
    [string]
    $Version
)

# download tool if necessary
if (-not(Get-Module ps2exe)) {
    Install-Module ps2exe
}

# build executable
ps2exe -inputFile ".\MMCSetupTool.ps1" -outputFile ".\bin\MMCSetupTool.exe" -iconFile ".\tools\mmc.ico" -title "MMC Setup Tool" -company "Multimedia Center Veenendaal" -copyright "© 2020 Bart Scholtus" -version "$Version" -requireAdmin 

# zip the files
Compress-Archive -Path ".\bin\MMCSetupTool.exe", ".\assets" -CompressionLevel Fastest -DestinationPath ".\bin\MMCSetupTool.zip"
