# this script needs Administrator rights, if not, it will stop
#Requires -RunAsAdministrator

# allow execution of this script, might prompt the user
Set-ExecutionPolicy RemoteSigned

# gives a user a choice: yes or no?
function Read-Choice (
   [Parameter(Mandatory)] [string] $Message,
   [Parameter()] [string[]] $Choices = ('&Ja','&Nee'),
   [Parameter()] [string] $DefaultChoice = '&Ja',
   [Parameter()] [string] $Question = "Ja of nee?"
) {
    $defaultIndex = $Choices.IndexOf($DefaultChoice)

    if ($defaultIndex -lt 0) {
        throw "$DefaultChoice is geen optie"
    }

    $choiceObj = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]

    foreach ($c in $Choices) {
        $choiceObj.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList $c))
    }

    $decision = $Host.UI.PromptForChoice($Message, $Question, $choiceObj, $defaultIndex)

    return $Choices[$decision]
}

# creates a folder for our assets
function Set-MMCFolder {
    # this is our desired location
    $loc = "$Env:USERPROFILE/Documents/MMC"

    # check if the folder already exists
    if ((Test-Path $loc)) {
        Write-Host -ForegroundColor Red "MMC-map bestaat al! Deze moet je eerst verwijderen voordat je deze tool opnieuw draait."
        
        # ask if the user wants to delete the folder
        $delete = Read-Choice -Message "Wil je dat ik de map voor je verwijder?"
        
        if ($delete -ne "&Ja") {
            Read-Host -Prompt "OK, dan moet je het zelf doen. Druk op ENTER om het programma te sluiten"
            exit
        }

        # delete it
        Remove-Item -Path "$Env:USERPROFILE/Documents/MMC" -Recurse

        Write-Host -ForegroundColor Green "It's gone! We gaan nu verder."
    }

    # create the folder
    New-Item -Path "$Env:USERPROFILE/Documents" -Name "MMC" -ItemType Directory | Out-Null
}

# sets the background of the user
function Install-Background {
    # our wallpaper
    $wallpaper = "assets/MMC Background.png"

    # create a new directory for our wallpaper
    New-Item -Path "$Env:windir\Web\Wallpaper" -Name "MMC" -ItemType Directory | Out-Null

    # copy the file
    Copy-Item $wallpaper "$Env:windir\Web\Wallpaper\MMC\"
    $wallpaper_file = "$Env:windir\Web\Wallpaper\MMC\MMC Background.png"

    # this is the key we need
    $regBG = "HKCU:\Control Panel\Desktop"

    # make registry changes
    Set-ItemProperty -Path $regBG -Name Wallpaper -Value $wallpaper_file
    Set-ItemProperty -Path $regBG -Name WallpaperStyle -Value "2"

    # refresh the wallpaper
    RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters 1, True

    Write-Host -ForegroundColor Green "Achtergrond is ingesteld!"
}

# sets the OEM info in the 'info' screen in Settings and on the System page in the Control Panel
function Set-OEMinfo {
    # this is our logo
    $logo = "assets/mmc.bmp"

    # copy the file
    Copy-Item $logo "$Env:windir\System32\"
    $logo_file = "$Env:windir\System32\mmc.bmp"

    # this is the key we need
    $regOEM = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation"
    
    # make registry changes
    Set-ItemProperty -Path $regOEM -Name Logo -Value $logo_file
    Set-ItemProperty -Path $regOEM -Name Manufacturer -Value "Multimedia Center Veenendaal"
    Set-ItemProperty -Path $regOEM -Name SupportPhone -Value "(+31) 0318 830 220"
    Set-ItemProperty -Path $regOEM -Name SupportHours -Value "09:30 - 17:30"
    Set-ItemProperty -Path $regOEM -Name SupportURL -Value "https://mmcveenendaal.nl"

    Write-Host -ForegroundColor Green "OEM info is ingesteld!"
}

# places some icons on the desktop
function Install-DesktopIcons {
    # the base registry key
    $reg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"

    # check if the key exists
    if (-not(Test-Path $reg)) {
        Write-Host -ForegroundColor Yellow "Register-entry bestaat nog niet, ff maken"

        # create the needed keys
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "HideDesktopIcons" | Out-Null
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons" -Name "NewStartPanel" | Out-Null
    }

    # This PC
    New-ItemProperty -Path $reg -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Value "0" -PropertyType DWORD -Force | Out-Null

    # User folder
    New-ItemProperty -Path $reg -Name "{59031a47-3f72-44a7-89c5-5595fe6b30ee}" -Value "0" -PropertyType DWORD -Force | Out-Null

    Write-Host -ForegroundColor Green "Snelkoppelingen zijn geplaatst! (ff scherm verversen)"
}

# set the proper start page of the Explorer
function Set-ThisPC {
    # the key we need
    $reg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

    # check if the key exists
    if (-not(Test-Path $reg)) {
        Write-Host -ForegroundColor Yellow "Register-entry bestaat nog niet, ff maken"

        # create the needed key
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "Advanced" | Out-Null
    }

    # This PC page
    New-ItemProperty -Path $reg -Name "LaunchTo" -Value "1" -PropertyType DWORD -Force | Out-Null

    Write-Host -ForegroundColor Green "Startpagina is ingesteld!"
}

# place instruction PDF on the desktop
function Install-InstructionPDF {
    # get our files
    $pdf = "assets/instructie.pdf"
    $icon = "assets/info.ico"

    # copy the files to the assets folder
    Copy-Item -Path $pdf -Destination "$Env:USERPROFILE/Documents/MMC"
    Copy-Item -Path $icon -Destination "$Env:USERPROFILE/Documents/MMC"

    # create a shortcut
    $shell = New-Object -comObject WScript.Shell
    $shortcut = $shell.CreateShortcut("$Env:USERPROFILE/Desktop/Welkom.lnk")
    $shortcut.TargetPath = "$Env:USERPROFILE/Documents/MMC/instructie.pdf"
    $shortcut.IconLocation = "$Env:USERPROFILE/Documents/MMC/info.ico"
    $shortcut.Save()

    Write-Host -ForegroundColor Green "Hij staat erop!"
}

# place our remote support link as icon on the desktop
function Install-RemoteSupport {
    # get our icon
    $icon = "assets/support.ico"

    # copy the icon to the assets folder
    Copy-Item -Path $icon -Destination "$Env:USERPROFILE/Documents/MMC"

    # create a shortcut
    $shell = New-Object -comObject WScript.Shell
    $shortcut = $shell.CreateShortcut("$Env:USERPROFILE/Desktop/Hulp op Afstand.lnk")
    $shortcut.TargetPath = "https://mmcveenendaal.nl/hulp-op-afstand"
    $shortcut.IconLocation = "$Env:USERPROFILE/Documents/MMC/support.ico"
    $shortcut.Save()

    Write-Host -ForegroundColor Green "Hij staat erop!"
}

# connect to our wifi network
function Connect-Wifi {
    # get the profile
    $wlanprofile = "assets/mmc_guest.xml"

    # install the profile
    netsh wlan add profile filename=$wlanprofile | Out-Null
    netsh wlan connect name="MMC_Guest" | Out-Null

    Write-Host -ForegroundColor Green "Verbonden met MMC_Guest!"
}

function Get-ActivationStatus {
    $status = Get-WmiObject SoftwareLicensingProduct -Filter "Name like 'Windows%'" | Where-Object { $_.LicenseStatus -eq 1 } | Select-Object -Property Description, LicenseStatus

    if ($status.LicenseStatus -eq 1) {
        Write-Host $status.Description
        Write-Host -ForegroundColor Green "Gefeliciteerd! Windows is geactiveerd."
    } else {
        Write-Host -ForegroundColor Red "LET OP: Windows is niet geactiveerd!"
    }
}

function Install-NewEdge {
    $url = "https://go.microsoft.com/fwlink/?linkid=2108834&Channel=Stable&language=nl"
    $file = "$Env:USERPROFILE/Documents/MMC/edge.exe"

    # download the installer
    Invoke-WebRequest -Uri $url -OutFile $file

    # run the installer
    $proc = Start-Process -FilePath $file -PassThru
    $proc.WaitForExit()

    Write-Host -ForegroundColor Green "Hij is geinstalleerd! (Hoop ik...)"
}

function Set-MicrosoftUpdateSetting {
    $reg = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

    if (-not(Test-Path $reg)) {
        Write-Host -ForegroundColor Yellow "Register-entry bestaat nog niet, ff maken"

        New-Item -Path "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate" -Name "AU" | Out-Null
    }

    New-ItemProperty -Path $reg -Name "AllowMUUpdateService" -Value "1" -PropertyType DWORD -Force | Out-Null

    Write-Host -ForegroundColor Green "Check!"
}

function Start-WindowsUpdate {
    # install NuGet
    Install-PackageProvider -Name "Nuget" -Force | Out-Null

    # install Windows Update module for PowerShell
    Install-Module PSWindowsUpdate -Force

    # get all available updates
    $updates = Get-WindowsUpdate -Install -Download

    if ($updates) {
        Write-Host -ForegroundColor Green "Er zijn updates! Ga ze ff installeren voor je, momentje"
        Write-Host -ForegroundColor Red "LET OP: er wordt automatisch opnieuw opgestart!!"

        Install-WindowsUpdate -AcceptAll -AutoReboot
    } else {
        Write-Host -ForegroundColor Red "Er zijn geen updates, woohoo!"
    }
}

$text = @"
+------------------------------------------------------------------------+
|             _ _ _ ____ _    _  _ ____ _  _    ___  _  _                |
|             | | | |___ |    |_/  |  | |\/|    |__] |  |                |
|             |_|_| |___ |___ | \_ |__| |  |    |__] | _|                |
|                                                                        |
|                            ___  ____                                   |
|                __ __ __    |  \ |___    __ __ __                       |
|                            |__/ |___                                   |
|                                                                        |
|    _  _ _  _ ____    ____ ____ ___ _  _ ___     ___ ____ ____ _        |
|    |\/| |\/| |       [__  |___  |  |  | |__]     |  |  | |  | |        |
|    |  | |  | |___    ___] |___  |  |__| |        |  |__| |__| |___     |
|                                                                        |
+------------------------------------------------------------------------+
"@

Write-Host $text -ForegroundColor Yellow

# CHECK FOR ASSETS
if (-not(Test-Path -Path "assets/")) {
    Write-Host -ForegroundColor Red "Assets ontbreken!"
    Read-Host -Prompt "Druk op ENTER om het programma te sluiten"
    exit
}

# SET ASSET FOLDER
Set-MMCFolder

# INSTALL BACKGROUND
$setupBG = Read-Choice -Message "Wil je de MMC-achtergrond instellen?"

if ($setupBG -eq "&Ja") {
    Install-Background
} else {
    Write-Host -ForegroundColor Yellow "Jammer... Hij is zo mooi!"
}

# SET OEM INFO
$setOEM = Read-Choice -Message "Wil je de OEM-info instellen?"

if ($setOEM -eq "&Ja") {
    Set-OEMinfo
} else {
    Write-Host -ForegroundColor Yellow "Prima, dan niet. Ook best."
}

# PLACE ICONS ON DESKTOP
$setupIcons = Read-Choice -Message "Wil je Deze pc en de gebruikersmap op het bureaublad plaatsen?"

if ($setupIcons -eq "&Ja") {
    Install-DesktopIcons
} else {
    Write-Host -ForegroundColor Yellow "Nou, dan niet he!"
}

# SET THIS PC AS START FOLDER
$setThisPC = Read-Choice -Message "Wil je Deze PC als startpagina instellen in de Verkenner?"

if ($setThisPC -eq "&Ja") {
    Set-ThisPC
} else {
    Write-Host -ForegroundColor Yellow "Weet waar je aan begint hoor..."
}

# INSTRUCTIEPDF OP BUREAUBLAD
$placeInstruction = Read-Choice -Message "Wil je de instructie-PDF op het bureaublad plaatsen?"

if ($placeInstruction -eq "&Ja") {
    Install-InstructionPDF
} else {
    Write-Host -ForegroundColor Yellow "Dan zullen ze er wel verstand van hebben..."
}

# HULP OP AFSTAND LINK OP BUREAUBLAD
$remoteSupport = Read-Choice -Message "Wil je de Hulp op Afstand-link op het bureaublad plaatsen?"

if ($remoteSupport -eq "&Ja") {
    Install-RemoteSupport
} else {
    Write-Host -ForegroundColor Yellow "Dan komen we wel langs ofzo als ze hulp nodig hebben"
}

# CONNECT TO WIFI
$connectWifi = Read-Choice -Message "Wil je met de wifi verbinden?"

if ($connectWifi -eq "&Ja") {
    Connect-Wifi
} else {
    Write-Host -ForegroundColor Yellow "Draadje dan maar?"
}

# CHECK FOR ACTIVATION
Write-Host -ForegroundColor Yellow "We gaan even de activatie van Windoosch checken"
Get-ActivationStatus

# INSTALL NEW EDGE
$installEdge = Read-Choice -Message "Wil je de nieuwe Edge installeren?"

if ($installEdge -eq "&Ja") {
    Install-NewEdge
} else {
    Write-Host -ForegroundColor Yellow "Zit het eindelijk standaard in Windows??"
}

# WINDOWS UPDATE SETTING ON
$setThisPC = Read-Choice -Message "Wil je updates voor Microsoft-producten inschakelen?"

if ($setThisPC -eq "&Ja") {
    Set-MicrosoftUpdateSetting
} else {
    Write-Host -ForegroundColor Yellow "Prima."
}

# RUN WINDOWS UPDATE
$setThisPC = Read-Choice -Message "Wil je Windows Update draaien?"

if ($setThisPC -eq "&Ja") {
    Start-WindowsUpdate
} else {
    Write-Host -ForegroundColor Yellow "Living on the edge?"
}

# time to finish things up

$text = @"
+-----------------------------------------+
|  ____ ____ ____ ____ ___  _ ____ ____   |
|  | __ |__/ |  | |___  |   | |___ [__    |
|  |__] |  \ |__| |___  |  _| |___ ___]   |
|                                         |
+-----------------------------------------+
"@

Write-Host $text -ForegroundColor Yellow
Write-Host -ForegroundColor Green "`nWe zijn wel zo'n beetje klaar, vergeet je de rest niet te doen??"
Read-Host -Prompt "Druk op ENTER om het programma te sluiten"
