using namespace System.Windows.Forms
using namespace System.Drawing

Add-Type -AssemblyName System.Windows.Forms

# Global variables
$OptionsView = New-Object ListView

$SetupToolForm = New-Object Form
$SetupToolForm.ClientSize = "500,350"
$SetupToolForm.text = "MMC Setup Tool - GUI"
$SetupToolForm.BackColor = "#ffffff"

function New-Label {
    Param(
        [Parameter(Mandatory)][string]$Text,
        [Parameter()][int]$Size = 9,
        [Parameter()][string]$Style = ""
    )

    $Label = New-Object Label
    $Label.Text = $Text
    $Label.AutoSize = $true
    $Label.Font = "Microsoft Sans Serif,$Size,$Style"

    $SetupToolForm.Controls.Add($Label)

    return $Label
}

function Show-Title {
    [void](New-Label -Text "MMC Setup Tool"-Size 18)
}

function Show-Update {
    $UpdateBtn = New-Label -Text "Update beschikbaar!" -Style "style=Bold,Underline"
    $UpdateBtn.Location = New-Object Point(200, 7)

    $UpdateBtn.Add_MouseEnter{ $this.Cursor = [Cursors]::Hand }
}

function Show-AdministratorStatus {
    $Title = New-Label -Text "Administratormodus: "
    $Title.Location = New-Object Point(0, 30)

    $Status = New-Label -Text "Ja"
    $Status.ForeColor = "Green"
    $Status.Location = New-Object Point(150, 30)
}

function Show-NetworkConnectionStatus {
    $Title = New-Label -Text "Internetverbinding: "
    $Title.Location = New-Object Point(250, 30)

    $Status = New-Label -Text "Ja"
    $Status.ForeColor = "Green"
    $Status.Location = New-Object Point(385, 30)
}

function Show-InstallationOptions {
    $OptionsView.Size = New-Object Size(500, 250)
    $OptionsView.Location = New-Object Point(0, 60)
    $OptionsView.CheckBoxes = $true
    $OptionsView.View = "Details"
    $OptionsView.AutoResizeColumns(1)

    $column = $OptionsView.Columns.Add("Selecteer welke onderdelen van toepassing zijn")
    $column.AutoResize(1)

    $defaultItems = @(
        "Windows activatiestatus controleren"
        "OEM-Informatie instellen"
        "Zet DezePC als startpagina in verkenner"
        "Plaats DezePC en de Gebruiker snelkoppeling op het bureaublad"
        "Plaats de MMC startpagina snelkoppeling op het bureaublad"
        "'mmcveenendaal.nl/start' instellen als startpagina in Edge"
        "Google instellen als standaard zoekmachine in Edge"
        "Voer de Windows Update uit"
        "Open de Windows Update (controlepunt)"
        "Voer voor Windows Store update(s) uit"
        "Open de Windows Store (controlepunt)"
        "Updates voor Windows producten inschakelen"
        "Open DeviceManager (controlepunt)"
        "Test de camera"
        "Test de speakers"
        "Hulp op afstand (Anydesk) installeren"
    )

    $optionalItems = @(
        "Microsoft 365 Personal / Family installeren"
        "Office 2021 Thuisgebruik en Studenten installeren"
        "Office 2021 Thuisgebruik en Zelfstandigen installeren"
        "G Data installeren"
    )

    $defaultItems | ForEach-Object {
        $item = [ListViewItem]::new($_)
        $item.Checked = $true
        [void]$OptionsView.Items.Add($item)
    }

    $optionalItems | ForEach-Object {
        $item = [ListViewItem]::new($_)
        [void]$OptionsView.Items.Add($item)
    }

    $SetupToolForm.Controls.Add($OptionsView)
}

function Start-Execute {
    # Execute checked option view items ...
}

function Show-Execute {
    $ExecuteBtn = New-Object Button
    $ExecuteBtn.Text = "Uitvoeren"
    $ExecuteBtn.Location = New-Object Point(10, 320)

    $ExecuteBtn.Add_MouseEnter{ $this.Cursor = [Cursors]::Hand }
    $ExecuteBtn.Add_Click{ Start-Execute }

    $SetupToolForm.Controls.Add($ExecuteBtn)
}

function Show-Status {
    $Title = New-Label -Text "Status: " -Style "style=Bold"
    $Title.Location = New-Object Point(100, 322)

    $Status = New-Label -Text "Inactief"
    $Status.Location = New-Object Point(150, 322)
}

Show-Title
Show-Update
Show-AdministratorStatus
Show-NetworkConnectionStatus
Show-InstallationOptions
Show-Execute
Show-Status

$SetupToolForm.ShowDialog()