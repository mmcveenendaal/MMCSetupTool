using namespace System.Windows.Forms
using namespace System.Drawing

Add-Type -AssemblyName System.Windows.Forms

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
    New-Label -Text "MMC Setup Tool"-Size 18
}

function Show-Update {
    $UpdateBtn = New-Label -Text "Update beschikbaar!" -Style "style=Bold,Underline"
    $UpdateBtn.Location = New-Object Point(200, 7)
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
    $OptionsView = New-Object ListView
    $OptionsView.Size = New-Object Size(500, 250)
    $OptionsView.Location = New-Object Point(0, 60)
    $OptionsView.CheckBoxes = $true
    $OptionsView.View = "Details"
    $OptionsView.AutoResizeColumns(1)

    $column = $OptionsView.Columns.Add("Selecteer welke onderdelen van toepassing zijn")
    $column.AutoResize(1)

    $items = @(
        [ListViewItem]::new("Windows activatiestatus controleren")
        [ListViewItem]::new("OEM-Informatie instellen")
        [ListViewItem]::new("Zet DezePC als startpagina in verkenner")
        [ListViewItem]::new("Plaats DezePC en de Gebruiker snelkoppeling op het bureaublad")
        [ListViewItem]::new("Plaats de MMC startpagina snelkoppeling op het bureaublad")
        [ListViewItem]::new("Voer de Windows Update uit")
        [ListViewItem]::new("Open de Windows Update (controlepunt)")
        [ListViewItem]::new("Voer voor Windows Store update(s) uit")
        [ListViewItem]::new("Open de Windows Store (controlepunt)")
        [ListViewItem]::new("Updates voor Windows producten inschakelen")
        [ListViewItem]::new("Open DeviceManager (controlepunt)")
        [ListViewItem]::new("Test de camera")
        [ListViewItem]::new("Test de speakers")
        [ListViewItem]::new("Microsoft 365 Personal / Family installeren")
        [ListViewItem]::new("Office 2021 Thuisgebruik en Studenten installeren")
        [ListViewItem]::new("Office 2021 Thuisgebruik en Zelfstandigen installeren")
        [ListViewItem]::new("G Data installeren")
        [ListViewItem]::new("Hulp op afstand (Anydesk) installeren")
    )

    foreach ($item in $items) {
        $item.Checked = $true
        $OptionsView.Items.Add($item) | Out-Null
    }

    $SetupToolForm.Controls.Add($OptionsView)
}

function Show-Execute {
    $ExecuteBtn = New-Object Button
    $ExecuteBtn.Text = "Uitvoeren"
    $ExecuteBtn.Location = New-Object Point(10, 320)

    $SetupToolForm.Controls.Add($ExecuteBtn)
}

function Show-Status {
    $Title = New-Label -Text "Status: " -Style "style=Bold"
    $Title.Location = New-Object Point(100, 322)

    $Status = New-Label -Text "In afwachting..."
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