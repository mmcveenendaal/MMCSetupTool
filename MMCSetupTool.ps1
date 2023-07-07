# some global vars
$version = '1.7'
$internet = $false

# check for admin rights
function Test-Administrator  {  
    Write-Host -ForegroundColor Cyan "`nControleren op Administrator rechten..."

    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    $admin = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

    if ($admin) {
        Write-Host -ForegroundColor Green "`tHet programma beschikt over Administrator rechten."
    } else {
        Write-Host -ForegroundColor Red "`tHet programma beschikt niet over Administrator rechten. Herstart het programma met Administrator rechten."
        Close-Program
    }
}

# close the program when the user hits ENTER
function Close-Program {
    Read-Host -Prompt "`nDruk op [ENTER] om het programma af te sluiten."
    exit
}

# get latest tool release
function Get-Update {
    Write-Host -ForegroundColor Cyan "`nDe Setup Tool laten controleren op updates..."

    # fix for Internet Explorer 'first run' error
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -Value 1

    # get the latest release
    $release = Invoke-WebRequest "https://github.com/matsn0w/MMCSetupTool/releases/latest" -Headers @{ "Accept" = "application/json" }
    $json = $release.Content | ConvertFrom-Json
    $latest = $json.tag_name

    # check if there is a newer version
    if ($latest -gt $version) {
        Write-Host -ForegroundColor Magenta "Er is een update beschrikbaar! $version => $latest"
                                               
        $update = Read-Choice -Message "Wil je de Setup Tool bijwerken?"

        if ($update -eq "&Ja") {
            Install-Update
        }
    } else {
        Write-Host -ForegroundColor Green "`tDe Setup Tool is up-to-date!"
    }
}

# self-update the tool
function Install-Update {
    Write-Host -ForegroundColor Cyan "`nDe Setup Tool updaten..."

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

    Write-Host -ForegroundColor Green "`tDe Setup Tool is geüpdatet!"

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
    # check for connection
    $status = Test-NetConnection

    return $status.PingSucceeded
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
    } else {
        Write-Host -ForegroundColor Yellow "`nGeen internet. De activatiestatus van Windows kan niet worden geverifiëerd."
    }

    # WINDOWS UPDATE SETTING ON
    Set-MicrosoftUpdateSetting

    # INSTALL WINDOWS STORE UPDATES
    if ($internet) {
        Update-Store
    } else {
        Write-Host -ForegroundColor Yellow "`nGeen internet. De Microsoft Store wordt niet bijgewerkt."
    }
}

# run program manual (ask for every function)
function Install-Manual {
    # SET OEM INFO
    $setOEM = Read-Choice -Message "`nWil je de OEM-info instellen?"

    if ($setOEM -eq "&Ja") {
        Set-OEMinfo
    }

    # PLACE ICONS ON DESKTOP
    $setupIcons = Read-Choice -Message "`nWil je Deze pc en de gebruikersmap op het bureaublad plaatsen?"

    if ($setupIcons -eq "&Ja") {
        Install-DesktopIcons
    }

    # SET THIS PC AS START FOLDER
    $setThisPC = Read-Choice -Message "`nWil je Deze PC als startpagina instellen in de Verkenner?"

    if ($setThisPC -eq "&Ja") {
        Set-ThisPC
    }

    # PLACE REMOTE SUPPORT ON DESKTOP
    $remoteSupport = Read-Choice -Message "`nWil je de MMC startpagina op het bureaublad plaatsen?"

    if ($remoteSupport -eq "&Ja") {
        Install-StartpageIcon
    }

    # SET DARK THEME
    $darkTheme = Read-Choice -Message "`nWil je de donkere modus inschakelen?"

    if ($darkTheme -eq "&Ja") {
        Set-DarkTheme
    }

    # CHECK FOR ACTIVATION
    if ($internet) {
        Get-ActivationStatus
    } else {
        Write-Host -ForegroundColor Yellow "`nGeen internet. De activatiestatus van Windows kan niet worden geverifiëerd."
    }

    # WINDOWS UPDATE SETTING ON
    $setThisPC = Read-Choice -Message "`nWil je updates voor Microsoft-producten inschakelen?"

    if ($setThisPC -eq "&Ja") {
        Set-MicrosoftUpdateSetting
    }

    # INSTALL WINDOWS STORE UPDATES
    if ($internet) {
        $updateStore = Read-Choice -Message "`nWil je de Windows Store bijwerken?"

        if ($updateStore -eq "&Ja") {
            Update-Store
        }
    } else {
        Write-Host -ForegroundColor Yellow "`nGeen internet. De Microsoft Store wordt niet bijgewerkt."
    }
}

# creates a folder for our assets
function Set-MMCFolder {
    # this is our desired location
    $loc = "$Env:USERPROFILE/Documents/MMC"

    # check if the folder already exists
    if ((Test-Path $loc)) {
        Write-Host -ForegroundColor Red "`tDe MMC-map bestaat al!"

        # delete it
        Remove-Item -Path "$Env:USERPROFILE/Documents/MMC" -Recurse

        Write-Host -ForegroundColor Green "`tDe MMC-map wordt opnieuw aangemaakt."
    }

    # create the folder
    New-Item -Path "$Env:USERPROFILE/Documents" -Name "MMC" -ItemType Directory | Out-Null
}

# sets the OEM info in the 'info' screen in Settings and on the System page in the Control Panel
function Set-OEMinfo {
    Write-Host -ForegroundColor Cyan "`nOEM instellen..."

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

    Write-Host -ForegroundColor Green "`tDe OEM is ingesteld!"
}

# places some icons on the desktop
function Install-DesktopIcons {
    Write-Host -ForegroundColor Cyan "`nIconen op het bureaublad plaatsen..."

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

    Write-Host -ForegroundColor Green "`tSnelkoppelingen zijn geplaatst (Het bureaublad moet mogelijk herladen worden)!"
}

# set the proper start page of the Explorer
function Set-ThisPC {
    Write-Host -ForegroundColor Cyan "`nStartpagina van de Verkenner instellen..."

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
    Write-Host -ForegroundColor Cyan "`nDe MMC startpagina op het bureaublad plaatsen..."

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

    Write-Host -ForegroundColor Green "`tDe MMC startpagina is op het bureaublad geplaatst."
}

# connect to our wifi network
function Connect-Wifi {
    Write-Host -ForegroundColor Cyan "`nHet apparaat met het MMC_Guest Wi-Fi verbinden..."

    # get the profile
    $wlanprofile = "assets/mmc_guest.xml"

    # install the profile
    netsh wlan add profile filename=$wlanprofile | Out-Null
    netsh wlan connect name="MMC_Guest" | Out-Null

    function Get-Connection {
        return Get-NetRoute | Where-Object DestinationPrefix -eq '0.0.0.0/0' | Get-NetIPInterface | Where-Object ConnectionState -eq 'Connected'
    }

    $connected = $false
    $count = 0

    while (!$connected) {
        if ($count++ -eq 5) {
            return $false
        }

        $connected = Get-Connection
        Start-Sleep -Seconds 3
    }

    return $true
}

# check the activation status of Windows
function Get-ActivationStatus {
    Write-Host -ForegroundColor Cyan "`nWindows activatiestatus controleren..."

    # get the status
    $status = Get-WmiObject SoftwareLicensingProduct -Filter "Name like 'Windows%'" | Where-Object { $_.LicenseStatus -eq 1 } | Select-Object -Property Description, LicenseStatus

    # check if Windows is active
    if ($status.LicenseStatus -eq 1) {
        Write-Host "`t$($status.Description)"
        Write-Host -ForegroundColor Green "`tWindows is geactiveerd."
    } else {
        Write-Host -ForegroundColor Red "`tLET OP: Windows is mogelijk niet geactiveerd!"
    }
}

# set the Microsoft products setting in Windows Update
function Set-MicrosoftUpdateSetting {
    Write-Host -ForegroundColor Cyan "`nUpdates voor Microsoft-producten inschakelen..."

    $ServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
    $ServiceManager.ClientApplicationID = "MMC Setup Tool"
    $ServiceManager.AddService2("7971f918-a847-4430-9279-4a52d1efe18d",7,"") | Out-Null

    Write-Host -ForegroundColor Green "`tUpdates zijn ingeschakeld."
}

# run Windows Update
function Start-WindowsUpdate {
    Write-Host -ForegroundColor Cyan "`nWindows Update starten..."

    # install NuGet
    Install-PackageProvider -Name "Nuget" -Force | Out-Null

    # install Windows Update module for PowerShell
    Install-Module PSWindowsUpdate -Force

    # get all available updates
    $updates = Get-WindowsUpdate

    # check if there are any updates
    if ($updates) {
        Write-Host -ForegroundColor Green "`tEr zijn updates gevonden die momenteel worden geïnstalleerd."

        # install the updates
        $updates | Install-WindowsUpdate -AcceptAll -Download -Install -IgnoreReboot
    } else {
        Write-Host -ForegroundColor Green "`tEr zijn geen updates gevonden."
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
    Write-Host -ForegroundColor Cyan "`nDrivers kunnenn niet automatisch worden geïnstalleerd... DeviceManager wordt geopend..."
    Start-Process devmgmt.msc

    # OPEN WINDOWS UPDATE
    Write-Host -ForegroundColor Cyan "`nWindows Update kan niet volledig werken op de automatische modus... Windows Update wordt geopend..."
    Start-Process ms-settings:windowsupdate
}

function Test-Hardware {
    Write-Host -ForegroundColor Cyan "`nWebcam wordt getest..."

    # check if there is a webcam installed
    $webcam = Get-CimInstance Win32_PnPEntity | Where-Object -Property PNPClass -eq 'Camera'

    if ($webcam) {
        # start the camera app to test the webcam
        explorer.exe shell:AppsFolder\Microsoft.WindowsCamera_8wekyb3d8bbwe!App
    } else {
        Write-Host -ForegroundColor Magenta "`tGeen webcam gevonden..."
    }

    Write-Host -ForegroundColor Cyan "`nSpeakers worden getest..."
    
    # test the speaker by playing some music
    $player = New-Object System.Media.SoundPlayer
    $player.SoundLocation = (Get-Location).Path + "/assets/audiotest.wav"
    $player.Play()
}

function Update-Store {
    Write-Host -ForegroundColor Cyan "`nWindows Store updates starten..."

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
            'Office 2021 Thuisgebruik en &Studenten',
            'Office 2021 Thuisgebruik en &Zelfstandigen'
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

    $version = Read-OfficeChoice -Message "`nKies een versie."

    switch ($version) {
        0 { $type = "O365HomePremRetail" }
        1 { $type = "HomeStudent2021Retail" }
        2 { $type = "HomeBusiness2021Retail" }

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

function Install-AnyDesk {
    $url = "https://get.anydesk.com/sMUo5Cx4/MMC_Hulp_op_Afstand.exe"
    $out = "$env:USERPROFILE/Desktop/MMC Hulp op Afstand.exe"

    Invoke-WebRequest -Uri $url -OutFile $out
}

# set window title
$Host.UI.RawUI.WindowTitle = "MMC Setup Tool v$version"

$text = @"
 __  __ __  __  ____   ____       _                 _____           _ 
|  \/  |  \/  |/ ___| / ___|  ___| |_ _   _ _ __   |_   _|__   ___ | |
| |\/| | |\/| | |     \___ \ / _ \ __| | | | '_ \    | |/ _ \ / _ \| |
| |  | | |  | | |___   ___) |  __/ |_| |_| | |_) |   | | (_) | (_) | |
|_|  |_|_|  |_|\____| |____/ \___|\__|\__,_| .__/    |_|\___/ \___/|_|
                                           |_|                               
"@

Write-Host -ForegroundColor Yellow $text

# CHECK FOR ADMIN RIGHTS
Test-Administrator

# ALLOW SCRIPT EXECUTION
Set-ExecutionPolicy RemoteSigned

# CHECK FOR ASSETS
if (-not(Test-Path -Path "assets/")) {
    Write-Host -ForegroundColor Red "`tAlle assets ontbreken. Herinstalleer de Setup Tool."
    Close-Program
}

# TEST INTERNET CONNECTION
Write-Host -ForegroundColor Cyan "`nInternetverbinding controlleren..."

$internet = Test-Internet

if ($internet) {
    Write-Host -ForegroundColor Green "`tHet apparaat is verbonden met het internet."
} else {
    Write-Host -ForegroundColor Red "`tHet apparaat is niet verbonden met het internet."
}

if (-not($internet)) {
    # CONNECT TO WIFI
    $connectWifi = Read-Choice -Message "`nWil je het apparaat met de MMC_Guest Wi-Fi laten verbinden?"

    if ($connectWifi -eq "&Ja") {
        $wifi = Connect-Wifi

        if ($wifi) {
            Write-Host -ForegroundColor Green "`tHet apparaat is nu verbonden op het MMC_Guest netwerk."
            $internet = Test-Internet
        } else {
            Write-Host -ForegroundColor Red "`tEr ging iets mis tijdens het verbinden... Verbind handmatig of gebruik een ethernet kabel."
        }
    }
}

if (-not($internet)) {
    Write-Host -ForegroundColor Yellow "`nLET OP: er is geen internetverbinding! De Setup Tool kan nu niet alle stappen uitvoeren. Verbind met een netwerk voor een optimale werking van de Setup Tool!"
}

# REMOVE OLD FILES
Remove-OldFilesAfterUpdate

# CHECK FOR UPDATES
if ($internet) {
    Get-Update
} else {
    Write-Host -ForegroundColor Yellow "`nGeen internetverbinding. Er wordt niet gecontroleerd op updates."
}

# SET ASSET FOLDER
Set-MMCFolder

# ASK FOR INSTALLATION TYPE
$type = Read-Choice -Message "`nWil je de Setup Tool automatisch laten uitvoeren?"

if ($type -eq "&Ja") {
    Install-Automatic
} else {
    Install-Manual
}

# OPEN WIDNOWS THAT NEED ATTENTION
Start-CheckWindows

# TEST SOME HARDWARE
Test-Hardware

# INSTALL ANYDESK
$installGDATA = Read-Choice "`nWil je het MMC Hulp op Afstand programma installeren?"

if ($installGDATA -eq "&Ja") {
    Install-AnyDesk
}

# INSTALL MICROSOFT OFFICE
$installOffice = Read-Choice -Message "`nWil je Office 365 installeren?"

if ($installOffice -eq "&Ja") {
    Install-Office
}

# INSTALL G DATA
$installGDATA = Read-Choice "`nWil je G DATA installeren?"

if ($installGDATA -eq "&Ja") {
    Install-GDATA
}

# RUN WINDOWS UPDATE
if ($internet) {
    $updateWindows = Read-Choice -Message "`nWil je de Windows Update installeren?"
    
    if ($updateWindows -eq "&Ja") {
        Start-WindowsUpdate
    }
} else {
    Write-Host -ForegroundColor Yellow "`nGeen internetverbinding. Windows Update wordt niet geïnstalleerd."
}

# time to finish things up
Write-Host -ForegroundColor Green "`nInstallatie afgerond!`nBedankt voor het gebruik maken van de MMC Setup Tool!"

# wait for the user to close the program
Close-Program
