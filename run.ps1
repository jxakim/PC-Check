[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$BIOSInfo = Get-WmiObject -Class Win32_BIOS

$icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Users\joa20072\OneDrive - Vestfold og Telemark fylkeskommune\Annet\Skrivebord\Fag Filer\Utvikling\Powershell Programmer\Bilder\PC.ico")

# FUNCTIONS
function Open_IPMenu {
    $InfoText1.text = "Nettverk Informasjon"
    $InfoText2.Visible = $False
    $Uptime_Text.Visible = $False
    $DiskSpace_Text.Visible = $False
    $programList.Visible = $False
    $ListPrograms_Button.Visible = $False
    $ClearPrograms_Button.Visible = $False
    $Username_Text.Visible = $False
    $RAMStorage_Text1.Visible = $False
    $RAMStorage_Text2.Visible = $False
    $Serial_Text.Visible = $False

    $IpAdress_Text.Visible = $True
    $MacAdress_Text.Visible = $True
}

function Open_SystemMenu {
    $InfoText1.text = "System Informasjon"
    $InfoText2.Visible = $False
    $IpAdress_Text.Visible = $False
    $MacAdress_Text.Visible = $False
    $RAMStorage_Text1.Visible = $False
    $RAMStorage_Text2.Visible = $False
    $Serial_Text.Visible = $False

    $Uptime_Text.Visible = $True
    $DiskSpace_Text.Visible = $True
    $Username_Text.Visible = $True
    $programList.Visible = $True
    $ListPrograms_Button.Visible = $True
    $ClearPrograms_Button.Visible = $True
}

function Open_BiosMenu {
    $InfoText1.text = "Bios Informasjon"
    $InfoText2.Visible = $False

    $Uptime_Text.Visible = $False
    $DiskSpace_Text.Visible = $False
    $Username_Text.Visible = $False
    $programList.Visible = $False
    $ListPrograms_Button.Visible = $False
    $ClearPrograms_Button.Visible = $False
    $IpAdress_Text.Visible = $False
    $MacAdress_Text.Visible = $False

    $RAMStorage_Text1.Visible = $True
    $RAMStorage_Text2.Visible = $True
    $Serial_Text.Visible = $True
}

# IP Meny
function Get_IPAdress {
    $result = ipconfig | Select-String -Pattern "IPv4 Address" | Select-Object -First 1 | ForEach-Object { ($_ -split ': ')[-1] }
    return $result
}

function Get_MACAdress {
    $result = (Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true })[0].MACAddress
    return $result
}

# System Meny
function List_Programs {
    $softwareList = Get_InstalledSoftware

    $programList.Items.Clear()

    foreach ($software in $softwareList) {
        $item = New-Object System.Windows.Forms.ListViewItem
        $item.Text = $software.Name

        $subItems = @()
        $subItems += $software.Version

        $item.SubItems.AddRange($subItems)

        $programList.Items.Add($item)
    }
}

function Get_InstalledSoftware {
    return Get-WmiObject -Class Win32_Product | Select-Object Name, Version
}

function Clear_Programs {
    $programList.Items.Clear()
}

function Get_Username {
    $result = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    return ($result -split '\\')[1]
}

function Get_Uptime {
    $uptime = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    $currentTime = Get-Date
    $uptimeDuration = $currentTime - $uptime
    $uptimeDurationFormatted = '{0} Dager, {1} Timer, {2} Minutter' -f $uptimeDuration.Days, $uptimeDuration.Hours, $uptimeDuration.Minutes

    if($uptimeDuration.Days -ge 1) {
        return "Datamaskinen din har vært på i " + $uptimeDurationFormatted + "`r`n" + "Det er lurt å ta en restart på datamaskinen!"
    } else {
        return "Datamaskinen din har vært på i " + $uptimeDurationFormatted
    }

}

function Get_DiskSpace {
    $FreeDiskSpace = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | Select-Object DeviceID, FreeSpace

    $FreeSpaceGB = [math]::Round($FreeDiskSpace.FreeSpace / 1GB, 2)

    return "$FreeSpaceGB GB ledig plass"
}

# Bios Meny

function Get_Bios {
    Write-Host "BIOS Version: $($BIOSInfo.SMBIOSBIOSVersion)"
    Write-Host "BIOS Release Date: $($BIOSInfo.ReleaseDate)"
    Write-Host "System Manufacturer: $($BIOSInfo.Manufacturer)"
    Write-Host "System Model: $($BIOSInfo.Caption)"
    Write-Host "BIOS Mode: $($BIOSInfo.BIOSCharacteristics)"
}

function Get_RAMStorage {
    param(
        $memory
    )

    $ram = Get-CimInstance -ClassName Win32_PhysicalMemory
    $totalMemory = ($ram | Measure-Object -Property Capacity -Sum).Sum / 1GB
    $availableMemory = (Get-CimInstance -ClassName Win32_OperatingSystem).FreePhysicalMemory / 1MB

    $MemoryInUse = [math]::Round($availableMemory, 2)
    $FreeMemory = [math]::Round(($totalMemory - $MemoryInUse), 2)

    if($memory -eq "inUse") {
        return "Du bruker $MemoryInUse gb av minnet ditt"
    } elseif($memory -eq "freeMemory") {
        return "Du har $FreeMemory gb ledig minne"
    }
}

function Get_Serial {
    $SerialNumber = (Get-WmiObject -Class "Win32_Bios").SerialNumber
    return "Serienummeret på pc'en din er: $SerialNumber"
}



# ########################################################## #
#                           Form Creation                    #
# ########################################################## #

$Form                                       = New-Object System.Windows.Forms.Form
$Form.Size                                  = New-Object System.Drawing.Size(800,500)
$Form.StartPosition                         = "CenterScreen"
$Form.FormBorderStyle                       = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$Form.Text                                  = "PC Sjekk Verktøy"
$Form.TopMost                               = $True
$Form.MaximizeBox                           = $false
$form.BackColor                             = [System.Drawing.ColorTranslator]::FromHtml("#e8e8e8")
$Form.Icon = $icon

# ########################################################## #
#                           Tekst                            #
# ########################################################## #

$InfoText1                                  = New-Object system.Windows.Forms.Label
$InfoText1.text                             = "Velg en av tjenestene på tabellen til venstre"
$InfoText1.AutoSize                         = $true
$InfoText1.Size                             = New-Object System.Drawing.Size(170, 30)
$InfoText1.location                         = New-Object System.Drawing.Point(250,25)
$InfoText1.Font                             = New-Object System.Drawing.Font('Microsoft Sans Serif',18,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$Form.Controls.Add($InfoText1)

$InfoText2                                  = New-Object system.Windows.Forms.Label
$InfoText2.text                             = "I dette programmet kan du se informasjon og hjelpemidler som kan hjelpe deg `r`nså slipper du kanskje å gå ned på IT-Servicedesken."
$InfoText2.AutoSize                         = $true
$InfoText2.Size                             = New-Object System.Drawing.Size(170, 30)
$InfoText2.location                         = New-Object System.Drawing.Point(260,70)
$InfoText2.Font                             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form.Controls.Add($InfoText2)

$Label1                                     = New-Object system.Windows.Forms.Label
$Label1.text                                = "Velg en tjeneste"
$Label1.AutoSize                            = $true
$Label1.Size                                = New-Object System.Drawing.Size(170, 30)
$Label1.location                            = New-Object System.Drawing.Point(25,25)
$Label1.Font                                = New-Object System.Drawing.Font('Microsoft Sans Serif',15,[System.Drawing.FontStyle]::Bold)
$Label1.BackColor                           = [System.Drawing.Color]::White
$Form.Controls.Add($Label1)

# Nettverk

$IpAdress_Text                              = New-Object System.Windows.Forms.Label
$IpAdress_Text.Text                         = "Din IP-Adresse er: $(Get_IPAdress)"
$IpAdress_Text.Size                         = New-Object System.Drawing.Size(500, 20)
$IpAdress_Text.Location                     = New-Object System.Drawing.Point(260, 70)
$IpAdress_Text.Font                         = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$IpAdress_Text.Visible                      = $False
$Form.Controls.Add($IpAdress_Text)

$MacAdress_Text                             = New-Object System.Windows.Forms.Label
$MacAdress_Text.Text                        = "Din MAC-Adresse er: $(Get_MACAdress)"
$MacAdress_Text.Size                        = New-Object System.Drawing.Size(500, 20)
$MacAdress_Text.Location                    = New-Object System.Drawing.Point(260, 100)
$MacAdress_Text.Font                        = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$MacAdress_Text.Visible                     = $False
$Form.Controls.Add($MacAdress_Text)


# System

$Uptime_Text                                = New-Object System.Windows.Forms.Label
$Uptime_Text.Text                           = $(Get_Uptime)
$Uptime_Text.Size                           = New-Object System.Drawing.Size(500, 30)
$Uptime_Text.Location                       = New-Object System.Drawing.Point(260, 70)
$Uptime_Text.Font                           = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$Uptime_Text.Visible                        = $False
$Form.Controls.Add($Uptime_Text)

$DiskSpace_Text                             = New-Object System.Windows.Forms.Label
$DiskSpace_Text.Text                        = "Disken din har $(Get_DiskSpace)"
$DiskSpace_Text.Size                        = New-Object System.Drawing.Size(500, 20)
$DiskSpace_Text.Location                    = New-Object System.Drawing.Point(260, 110)
$DiskSpace_Text.Font                        = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$DiskSpace_Text.Visible                     = $False
$Form.Controls.Add($DiskSpace_Text)

$Username_Text                             = New-Object System.Windows.Forms.Label
$Username_Text.Text                        = "Ditt brukernavn er: $(Get_Username)"
$Username_Text.Size                        = New-Object System.Drawing.Size(500, 20)
$Username_Text.Location                    = New-Object System.Drawing.Point(260, 140)
$Username_Text.Font                        = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$Username_Text.Visible                     = $False
$Form.Controls.Add($Username_Text)

$programList                                = New-Object System.Windows.Forms.ListView
$programList.Size                           = New-Object System.Drawing.Size(500, 205)
$programList.Location                       = New-Object System.Drawing.Point(260, 230)
$programList.Font                           = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$programList.View                           = [System.Windows.Forms.View]::Details
$programList.Visible                        = $False

$programList.Columns.Add("Programnavn", 350)
$programList.Columns.Add("Versjon", 200)

$Form.Controls.Add($programList)

# Bios

$RAMStorage_Text1                            = New-Object System.Windows.Forms.Label
$RAMStorage_Text1.Text                       = "$(Get_RAMStorage "freeMemory")"
$RAMStorage_Text1.Size                       = New-Object System.Drawing.Size(500, 30)
$RAMStorage_Text1.Location                   = New-Object System.Drawing.Point(260, 70)
$RAMStorage_Text1.Font                       = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$RAMStorage_Text1.Visible                    = $False
$Form.Controls.Add($RAMStorage_Text1)

$RAMStorage_Text2                            = New-Object System.Windows.Forms.Label
$RAMStorage_Text2.Text                       = "$(Get_RAMStorage "inUse")"
$RAMStorage_Text2.Size                       = New-Object System.Drawing.Size(500, 30)
$RAMStorage_Text2.Location                   = New-Object System.Drawing.Point(260, 100)
$RAMStorage_Text2.Font                       = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$RAMStorage_Text2.Visible                    = $False
$Form.Controls.Add($RAMStorage_Text2)

$Serial_Text                            = New-Object System.Windows.Forms.Label
$Serial_Text.Text                       = "$(Get_Serial)"
$Serial_Text.Size                       = New-Object System.Drawing.Size(500, 30)
$Serial_Text.Location                   = New-Object System.Drawing.Point(260, 130)
$Serial_Text.Font                       = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$Serial_Text.Visible                    = $False
$Form.Controls.Add($Serial_Text)


# ########################################################## #
#                           KNAPPER                          #
# ########################################################## #

$Nettverk_Button                            = New-Object System.Windows.Forms.Button
$Nettverk_Button.Text                       = "Nettverk Informasjon"
$Nettverk_Button.Size                       = New-Object System.Drawing.Size(170, 40)
$Nettverk_Button.Location                   = New-Object System.Drawing.Point(25, 70)
$Nettverk_Button.Font                       = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$Nettverk_Button.Add_Click({Open_IPMenu})
$Form.Controls.Add($Nettverk_Button)

$System_Button                              = New-Object System.Windows.Forms.Button
$System_Button.Text                         = "System Informasjon"
$System_Button.Size                         = New-Object System.Drawing.Size(170, 40)
$System_Button.Location                     = New-Object System.Drawing.Point(25, 130)
$System_Button.Font                         = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$System_Button.Add_Click({Open_SystemMenu})
$Form.Controls.Add($System_Button)

$Bios_Button                                = New-Object System.Windows.Forms.Button
$Bios_Button.Text                           = "Bios Informasjon"
$Bios_Button.Size                           = New-Object System.Drawing.Size(170, 40)
$Bios_Button.Location                       = New-Object System.Drawing.Point(25, 190)
$Bios_Button.Font                           = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$Bios_Button.Add_Click({Open_BiosMenu})
$Form.Controls.Add($Bios_Button)

$ListPrograms_Button                        = New-Object System.Windows.Forms.Button
$ListPrograms_Button.Text                   = "Se programmer på enhet"
$ListPrograms_Button.Size                   = New-Object System.Drawing.Size(199, 30)
$ListPrograms_Button.Location               = New-Object System.Drawing.Point(260, 190)
$ListPrograms_Button.BackColor              = [System.Drawing.Color]::LightGreen
$ListPrograms_Button.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$ListPrograms_Button.Visible                = $False
$ListPrograms_Button.Add_Click({List_Programs})
$Form.Controls.Add($ListPrograms_Button)

$ClearPrograms_Button                        = New-Object System.Windows.Forms.Button
$ClearPrograms_Button.Text                   = "Clear"
$ClearPrograms_Button.Size                   = New-Object System.Drawing.Size(99, 30)
$ClearPrograms_Button.Location               = New-Object System.Drawing.Point(461, 190)
$ClearPrograms_Button.BackColor              = [System.Drawing.Color]::Red
$ClearPrograms_Button.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$ClearPrograms_Button.Visible                = $False
$ClearPrograms_Button.Add_Click({Clear_Programs})
$Form.Controls.Add($ClearPrograms_Button)


# ########################################################## #
#                           Annet                            #
# ########################################################## #

$Panel1                                     = New-Object system.Windows.Forms.Panel
$Panel1.Size                                = New-Object System.Drawing.Size(220, 460)
$Panel1.BackColor                           = [System.Drawing.Color]::White
$Form.Controls.Add($Panel1)

[void] $Form.Add_Shown($Form.Activate())
[void] $Form.ShowDialog() 