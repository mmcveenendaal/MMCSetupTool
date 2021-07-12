# some global vars
$version = '1.6.2'
$Global:internet = $false

# check for admin rights
function Test-Administrator  {  
    Write-Host -ForegroundColor Yellow "`nEven kijken of je admin bent..."

    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    $admin = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

    if ($admin) {
        Write-Host -ForegroundColor Green "`tHelemaal goed!"
    } else {
        Write-Host -ForegroundColor Red "`tStart ff als admin jongeh! Zo ken ik toch niks??"
        Close-Program
    }
}

# close the program when the user hits ENTER
function Close-Program {
    Read-Host -Prompt "`nDruk op ENTER om het programma te sluiten"
    exit
}

# get latest tool release
function Get-Update {
    Write-Host -ForegroundColor Yellow "`nUpdates ophalen..."

    # fix for Internet Explorer 'first run' error
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -Value 1

    # get the latest release
    $release = Invoke-WebRequest "https://github.com/matsn0w/MMCSetupTool/releases/latest" -Headers @{ "Accept" = "application/json" }
    $json = $release.Content | ConvertFrom-Json
    $latest = $json.tag_name

    # check if there is a newer version
    if ($latest -gt $version) {
        Write-Host -ForegroundColor Magenta "Er is een update beschikbaar!`nHuidige versie: $version`nNieuwste versie: $latest"
                                               
        $update = Read-Choice -Message "Wil je de tool bijwerken?"

        if ($update -eq "&Ja") {
            Install-Update
        }
    } else {
        Write-Host -ForegroundColor Green "`tDe tool is up-to-date!"
    }
}

# self-update the tool
function Install-Update {
    Write-Host -ForegroundColor Yellow "`nUpdate installeren..."

    # get the latest release (dynamic url)
    $url = "https://github.com/matsn0w/MMCSetupTool/releases/latest/download/MMCSetupTool.zip"

    # download the zip
    Invoke-WebRequest -Uri $url -OutFile "assets/update.zip"

    # extract the zip
    Expand-Archive -Path "assets/update.zip" -DestinationPath "update"

    # delete the zip
    Remove-Item -Path "assets/update.zip"

    # rename current exe
    Rename-Item "MMCSetupTool.exe" -NewName "MMCSetupTool.exe.tmp"

    # delete old stuff
    Remove-Item ".\assets" -Recurse

    # move new files
    Get-ChildItem -Path ".\update" | Move-Item -Destination "." -Force

    # delete update folder
    Remove-Item ".\update" -Recurse

    Write-Host -ForegroundColor Green "`tDe tool is bijgewerkt!"

    # start the new tool
    Start-Process "MMCSetupTool.exe" -Verb RunAs

    # exit the old tool
    exit
}

# check if old exe exists
function Remove-OldFilesAfterUpdate {
    $file = "MMCSetupTool.exe.tmp"

    if (Test-Path $file) {
        # delete the file
        Remove-Item $file -Force
    }
}

# tests the internet connection
function Test-Internet {
    Write-Host -ForegroundColor Yellow "`nInternetverbinding testen..."

    # check for connection
    $status = Test-NetConnection "google.com"

    if ($status.PingSucceeded) {
        Write-Host -ForegroundColor Green "`tEr is een internetverbinding!"
        return $true
    }

    Write-Host -ForegroundColor Red "`tEr is geen internetverbinding!"
    return $false
}

# gives a user a choice: yes or no?
function Read-Choice (
   [Parameter(Mandatory)] [string] $Message,
   [Parameter()] [string[]] $Choices = ('&Ja','&Nee'),
   [Parameter()] [string] $DefaultChoice = '&Ja',
   [Parameter()] [string] $Question = "Ja of nee?"
) {
    $defaultIndex = $Choices.IndexOf($DefaultChoice)

    $choiceObj = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]

    foreach ($c in $Choices) {
        $choiceObj.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList $c))
    }

    $decision = $Host.UI.PromptForChoice($Message, $Question, $choiceObj, $defaultIndex)

    return $Choices[$decision]
}

# run all functions automatically
function Install-Automatic {
    # SET OEM INFO
    Set-OEMinfo

    # PLACE ICONS ON DESKTOP
    Install-DesktopIcons

    # SET THIS PC AS START FOLDER
    Set-ThisPC

    # PLACE REMOTE SUPPORT ON DESKTOP
    Install-StartpageIcon

    # SET DARK THEME
    Set-DarkTheme

    # CHECK FOR ACTIVATION
    if ($internet) {
        Get-ActivationStatus
    }

    # WINDOWS UPDATE SETTING ON
    Set-MicrosoftUpdateSetting

    # RUN WINDOWS UPDATE
    if ($internet) {        
        Start-WindowsUpdate
    }    
}

# run program manual (ask for every function)
function Install-Manual {
    # SET OEM INFO
    $setOEM = Read-Choice -Message "`nWil je de OEM-info instellen?"

    if ($setOEM -eq "&Ja") {
        Set-OEMinfo
    } else {
        Write-Host -ForegroundColor Cyan "`tPrima, dan niet. Ook best."
    }

    # PLACE ICONS ON DESKTOP
    $setupIcons = Read-Choice -Message "`nWil je Deze pc en de gebruikersmap op het bureaublad plaatsen?"

    if ($setupIcons -eq "&Ja") {
        Install-DesktopIcons
    } else {
        Write-Host -ForegroundColor Cyan "`tNou, dan niet he!"
    }

    # SET THIS PC AS START FOLDER
    $setThisPC = Read-Choice -Message "`nWil je Deze PC als startpagina instellen in de Verkenner?"

    if ($setThisPC -eq "&Ja") {
        Set-ThisPC
    } else {
        Write-Host -ForegroundColor Cyan "`tWeet waar je aan begint hoor..."
    }

    # PLACE REMOTE SUPPORT ON DESKTOP
    $remoteSupport = Read-Choice -Message "`nWil je de MMC Start-link op het bureaublad plaatsen?"

    if ($remoteSupport -eq "&Ja") {
        Install-StartpageIcon
    } else {
        Write-Host -ForegroundColor Cyan "`tDan komen we wel langs of zo als ze hulp nodig hebben"
    }

    # SET DARK THEME
    $darkTheme = Read-Choice -Message "`nWil je het donkere thema instellen?"

    if ($darkTheme -eq "&Ja") {
        Set-DarkTheme
    } else {
        Write-Host -ForegroundColor Cyan "`tShine bright like a diamond!"
    }

    # CHECK FOR ACTIVATION
    if ($internet) {
        Get-ActivationStatus
    }

    # WINDOWS UPDATE SETTING ON
    $setThisPC = Read-Choice -Message "`nWil je updates voor Microsoft-producten inschakelen?"

    if ($setThisPC -eq "&Ja") {
        Set-MicrosoftUpdateSetting
    } else {
        Write-Host -ForegroundColor Cyan "`tPrima."
    }

    # RUN WINDOWS UPDATE
    if ($internet) {
        $setThisPC = Read-Choice -Message "`nWil je Windows Update draaien?"
        
        if ($setThisPC -eq "&Ja") {
            Start-WindowsUpdate
        } else {
            Write-Host -ForegroundColor Cyan "`tLiving on the edge?"
        }
    }
}

# creates a folder for our assets
function Set-MMCFolder {
    # this is our desired location
    $loc = "$Env:USERPROFILE/Documents/MMC"

    # check if the folder already exists
    if ((Test-Path $loc)) {
        Write-Host -ForegroundColor Red "`tMMC-map bestaat al!"

        # delete it
        Remove-Item -Path "$Env:USERPROFILE/Documents/MMC" -Recurse

        Write-Host -ForegroundColor Green "`tMap verwijderd! We gaan nu verder. Ik maak 'm opnieuw aan voor je."
    }

    # create the folder
    New-Item -Path "$Env:USERPROFILE/Documents" -Name "MMC" -ItemType Directory | Out-Null
}

# sets the OEM info in the 'info' screen in Settings and on the System page in the Control Panel
function Set-OEMinfo {
    Write-Host -ForegroundColor Yellow "`nOEM info instellen..."

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

    Write-Host -ForegroundColor Green "`tOEM info is ingesteld!"
}

# places some icons on the desktop
function Install-DesktopIcons {
    Write-Host -ForegroundColor Yellow "`nIcoontjes op bureaublad plaatsen..."

    # the base registry key
    $reg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"

    # check if the key exists
    if (-not(Test-Path $reg)) {
        # create the needed keys
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "HideDesktopIcons" | Out-Null
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons" -Name "NewStartPanel" | Out-Null
    }

    # This PC
    New-ItemProperty -Path $reg -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Value "0" -PropertyType DWORD -Force | Out-Null

    # User folder
    New-ItemProperty -Path $reg -Name "{59031a47-3f72-44a7-89c5-5595fe6b30ee}" -Value "0" -PropertyType DWORD -Force | Out-Null

    Write-Host -ForegroundColor Green "`tSnelkoppelingen zijn geplaatst! (mogelijk ff scherm verversen)"
}

# set the proper start page of the Explorer
function Set-ThisPC {
    Write-Host -ForegroundColor Yellow "`nStartpagina van de Verkenner instellen..."

    # the key we need
    $reg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

    # check if the key exists
    if (-not(Test-Path $reg)) {
        # create the needed key
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "Advanced" | Out-Null
    }

    # This PC page
    New-ItemProperty -Path $reg -Name "LaunchTo" -Value "1" -PropertyType DWORD -Force | Out-Null

    Write-Host -ForegroundColor Green "`tStartpagina is ingesteld!"
}

# place our remote support link as icon on the desktop
function Install-StartpageIcon {
    Write-Host -ForegroundColor Yellow "`nStartpaginalinkje op bureaublad plaatsen..."

    # get our icon
    $icon = "assets/mmc.ico"

    # copy the icon to the assets folder
    Copy-Item -Path $icon -Destination "$Env:USERPROFILE/Documents/MMC"

    # create a shortcut
    $shell = New-Object -comObject WScript.Shell
    $shortcut = $shell.CreateShortcut("$Env:USERPROFILE/Desktop/MMC Start.lnk")
    $shortcut.TargetPath = "https://mmcveenendaal.nl/start"
    $shortcut.IconLocation = "$Env:USERPROFILE/Documents/MMC/mmc.ico"
    $shortcut.Save()

    Write-Host -ForegroundColor Green "`tHij staat erop!"
}

# connect to our wifi network
function Connect-Wifi {
    Write-Host -ForegroundColor Yellow "`nMet wifi verbinden..."

    # get the profile
    $wlanprofile = "assets/mmc_guest.xml"

    # install the profile
    netsh wlan add profile filename=$wlanprofile | Out-Null
    netsh wlan connect name="MMC_Guest" | Out-Null

    Write-Host -ForegroundColor Green "`tVerbonden met MMC_Guest!"
}

# check the activation status of Windows
function Get-ActivationStatus {
    Write-Host -ForegroundColor Yellow "`nWe gaan even kijken of Windoosch is geactiveerd..."

    # get the status
    $status = Get-WmiObject SoftwareLicensingProduct -Filter "Name like 'Windows%'" | Where-Object { $_.LicenseStatus -eq 1 } | Select-Object -Property Description, LicenseStatus

    # check if Windows is active
    if ($status.LicenseStatus -eq 1) {
        Write-Host "`t$($status.Description)"
        Write-Host -ForegroundColor Green "`tGefeliciteerd! Windows is geactiveerd."
    } else {
        Write-Host -ForegroundColor Red "`tLET OP: Windows is mogelijk niet geactiveerd!"
    }
}

# set the Microsoft products setting in Windows Update
function Set-MicrosoftUpdateSetting {
    Write-Host -ForegroundColor Yellow "`nUpdates voor Microsoft-producten inschakelen..."

    $ServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
    $ServiceManager.ClientApplicationID = "MMC Setup Tool"
    $ServiceManager.AddService2("7971f918-a847-4430-9279-4a52d1efe18d",7,"") | Out-Null

    Write-Host -ForegroundColor Green "`tCheck!"
}

# run Windows Update
function Start-WindowsUpdate {
    Write-Host -ForegroundColor Yellow "`nWindows Update starten..."

    # install NuGet
    Install-PackageProvider -Name "Nuget" -Force | Out-Null

    # install Windows Update module for PowerShell
    Install-Module PSWindowsUpdate -Force

    # get all available updates
    $updates = Get-WindowsUpdate -Install -Download

    # check if there are any updates
    if ($updates) {
        Write-Host -ForegroundColor Green "`tEr zijn updates! Ga ze ff installeren voor je, momentje"

        # install the updates
        Install-WindowsUpdate -AcceptAll
    } else {
        Write-Host -ForegroundColor Green "`tEr zijn geen updates, woohoo!"
    }
}

function Set-DarkTheme {   
    # set app theme to dark 
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value "0"
    
    # set system theme to dark
    Set-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value "0"
}

function Start-CheckWindows {
    # OPEN DEVICE MANAGER
    Write-Host -ForegroundColor Yellow "`nControleer je efskes of alle drivers zijn geïnstalleerd??"
    Start-Process devmgmt.msc

    Write-Host -ForegroundColor Green "`tDevice Manager is voor je geopend"
    
    # OPEN WINDOWS UPDATE
    Write-Host -ForegroundColor Yellow "`nIk ben niet zo goed met die updates, dus hou even een oogje in het zeil..."
    Start-Process ms-settings:windowsupdate
    
    Write-Host -ForegroundColor Green "`tWindows Update is voor je geopend"
}

function Test-Hardware {
    Write-Host -ForegroundColor Yellow "`nWe testen even de webcam."

    # check if there is a webcam installed
    $webcam = Get-CimInstance Win32_PnPEntity | Where-Object -Property PNPClass -eq 'Camera'

    if ($webcam) {
        Write-Host -ForegroundColor Cyan "`nLachen!"

        # start the camera app to test the webcam
        explorer.exe shell:AppsFolder\Microsoft.WindowsCamera_8wekyb3d8bbwe!App
    } else {
        Write-Host -ForegroundColor Cyan "`nGeen webcam gevonden..."
    }

    Write-Host -ForegroundColor Yellow "`nZou de speaker het doen?"
    
    # test the speaker by playing some music
    $player = New-Object System.Media.SoundPlayer
    $player.SoundLocation = (Get-Location).Path + "/assets/audiotest.wav"
    $player.Play()

    Write-Host -ForegroundColor Green "`tIk heb zelf geen idee of 't werkte allemaal, maar dat zoek je zelf maar uit"
}

function Update-Store {
    Write-Host -ForegroundColor Yellow "`nWindows Store updates starten..."

    # start updates
    Get-CimInstance -Namespace "root\cimv2\mdm\dmmap" -ClassName "MDM_EnterpriseModernAppManagement_AppManagement01" | 
    Invoke-CimMethod -MethodName UpdateScanMethod | Out-Null

    # launch Updates page in Windows Store app
    Start-Process ms-windows-store:updates

    Write-Host -ForegroundColor Green "`tUpdates zijn gestart!"
}

function Install-Office {
    function Read-OfficeChoice (
        [Parameter(Mandatory)] [string] $Message,
        [Parameter()] [string[]] $Choices = (
            '&Microsoft 365 Personal / Family',
            'Office 2019 Thuisgebruik en &Studenten',
            'Office 2019 Thuisgebruik en &Zelfstandigen'
        ),
        [Parameter()] [string] $DefaultChoice = 0, # 365
        [Parameter()] [string] $Question = "Selecteer een versie"
     ) {
         $choiceObj = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
     
         foreach ($c in $Choices) {
             $choiceObj.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList $c))
         }
     
         $decision = $Host.UI.PromptForChoice($Message, $Question, $choiceObj, $DefaultChoice)
     
         return $decision
     }

    $version = Read-OfficeChoice -Message "`nWelke versie gaat 't worden?"

    switch ($version) {
        0 { $type = "O365HomePremRetail" }
        1 { $type = "HomeStudent2019Retail" }
        2 { $type = "HomeBusiness2019Retail" }

        Default {}
    }

    $url = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=$type&platform=x64&language=nl-nl"
    $out = "$env:USERPROFILE/Downloads/Office_Setup.exe"

    Invoke-WebRequest -Uri $url -OutFile $out
    Start-Process -FilePath $out
}

function Install-GDATA {
    $url = "https://gdata-a.akamaihd.net/Q/WEB/B2C/WEU/GDATA_INTERNETSECURITY_WEB_WEU.exe"
    $out = "$env:USERPROFILE/Downloads/GDATA_Setup.exe"

    Invoke-WebRequest -Uri $url -OutFile $out
    Start-Process -FilePath $out
}

# set window title
$Host.UI.RawUI.WindowTitle = "MMC Setup Tool v$version"

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

Write-Host -ForegroundColor Yellow $text

# CHECK FOR ADMIN RIGHTS
Test-Administrator

# ALLOW SCRIPT EXECUTION
Set-ExecutionPolicy RemoteSigned

# CHECK FOR ASSETS
if (-not(Test-Path -Path "assets/")) {
    Write-Host -ForegroundColor Red "`tAssets ontbreken! Download de tool opnieuw."
    Close-Program
}

# TEST INTERNET CONNECTION
$internet = Test-Internet

if (-not($internet)) {
    # CONNECT TO WIFI
    $connectWifi = Read-Choice -Message "Wil je met de wifi verbinden?"

    if ($connectWifi -eq "&Ja") {
        Connect-Wifi
    } else {
        Write-Host -ForegroundColor Cyan "`tDraadje dan maar?"
    }
}

# REMOVE OLD FILES
Remove-OldFilesAfterUpdate

# CHECK FOR UPDATES
if ($internet) {
    Get-Update
}

# SET ASSET FOLDER
Set-MMCFolder

# ASK FOR INSTALLATION TYPE
$type = Read-Choice -Message "`nWil je de tool automatisch uitvoeren?"

if ($type -eq "&Ja") {
    Write-Host -ForegroundColor Yellow "`tDe rest gaat vanzelf, hou je vast!"
    Install-Automatic
} else {
    Write-Host -ForegroundColor Yellow "`tJe wordt bij elke functie gevraagd wat er moet gebeuren."
    Install-Manual
}

# OPEN WIDNOWS THAT NEED ATTENTION
Start-CheckWindows

# TEST SOME HARDWARE
Test-Hardware

# INSTALL WINDOWS STORE UPDATES
Update-Store

# INSTALL MICROSOFT OFFICE
$installOffice = Read-Choice -Message "`nWil je Office installeren?"

if ($installOffice -eq "&Ja") {
    Install-Office
} else {
    Write-Host -ForegroundColor Cyan "`tLekker Live Mail gebruiken lol"
}

# INSTALL G DATA
$installGDATA = Read-Choice "`nWil je G DATA installeren?"

if ($installGDATA -eq "&Ja") {
    Install-GDATA
} else {
    Write-Host -ForegroundColor Cyan "`tBetter safe than sorry..."
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

if (-not($internet)) {
    Write-Host -ForegroundColor Cyan "`tOh ja, er was geen internet dus ik heb niet alles gedaan... Check je dat nog ff?"
}

# wait for the user to close the program
Close-Program
