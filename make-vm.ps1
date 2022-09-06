<#
This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License
as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty
of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
You should have received a copy of the GNU General Public License along with this program. If not, see <https://www.gnu.org/licenses/>.
#>

$session = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (!$session.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host -ForegroundColor Red "This script requires administrative privileges.`nExiting."
    return
}

# GUI to choose base path
function PickFolder {
    param (
        [Parameter()]
        [string] $Object
    )

    # Access built-in library
    Add-Type -AssemblyName System.Windows.Forms

    # Initialize pop-up
    $picker = New-Object $Object
    $ps_version = $PSVersionTable.PSVersion | Select-Object -ExpandProperty Major
    # Powershell versioning shim
    if ($ps_version -gt "5") {
        $picker.InitialDirectory = "C:\"
    } else {
        $picker.RootFolder = "MyComputer"
        if ($Object -eq "System.Windows.Forms.FolderBrowserDialog") {
            $picker.SelectedPath = "C:\"
        }
    }
    # Only allow files of type ISO when picking files
    if ($Object -eq "System.Windows.Forms.OpenFileDialog") {
        $picker.Filter = "iso files (*.iso)|*.iso"
    }

    # Call and save choice
    if ($picker.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($Object -eq "System.Windows.Forms.OpenFileDialog") {
            return $picker.FileName
        } else {
            return $picker.SelectedPath
        }
    }
}

# GUI to choose virtual switch
function PickSwitch {
    # Access built-in library
    Add-Type -AssemblyName System.Windows.Forms

    # Parent window
    $select = New-Object system.Windows.Forms.Form
    $select.ClientSize = '400,240'
    $select.text = "Switch picker"
    $select.BackColor = '#ffffff'

    # Combo box
    $list = New-Object system.Windows.Forms.ComboBox
    $list.text = ""
    $list.width = 300
    $list.location = New-Object System.Drawing.Point(50,100)
    $list.Font = 'Microsoft Sans Serif, 12'
    Get-VMSwitch | Select-Object -ExpandProperty Name | ForEach-Object {[void] $list.Items.Add($_)}

    # Label
    $label = New-Object system.Windows.Forms.Label
    $label.text = "Select a virtual switch:"
    $label.width = (400-2*47)
    $label.location = New-Object System.Drawing.Point(47,60)
    $label.Font = 'Microsoft Sans Serif, 12'

    # Confirmation button
    $button = New-Object System.Windows.Forms.Button
    $button.text = "OK"
    $button.width = 100
    $button.location = New-Object System.Drawing.Point(250,180)
    $button.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $select.AcceptButton = $button

    # Attach everything to parent window
    $select.Controls.Add($list)
    $select.Controls.Add($label)
    $select.Controls.Add($button)

    # Initialize
    $select.Topmost = $true
    $select_result = $select.ShowDialog()

    # Save choice
    if ($select_result -eq [System.Windows.Forms.DialogResult]::OK) {
        $vm_switch_name = $list.SelectedItem
        return $vm_switch_name
    }
}

# Define colors
$verbose_color = "White"
$success_color = "Green"

Write-Host "`n`n    Create a virtual machine for use with Hyper-V`n    Defaults can be accepted by pressing [ENTER]`n    Copyright 2022, Marian Arlt, All rights reserved`n`n"

# Create parameter array
$parameters = @{}

# Prompt for config options
$accept_defaults = Read-Host "Do you want to skip configurations that have default values? ([Y]es/[N]o Default: N)"
if ($accept_defaults -like "N*") {
    $accept_defaults = $false
}

while (!$vm_name) {
    $vm_name = Read-Host "Name your machine"
}
$parameters.Add("Name", $vm_name)
Write-Host -ForegroundColor $verbose_color "> The new machine will be called $vm_name."

$vm_base_path = Read-Host "`nPress [ENTER] to choose the base folder for the VM or type a path"
if (!$vm_base_path) {
    $vm_base_path = PickFolder -Object System.Windows.Forms.FolderBrowserDialog
}
$parameters.Add("Path", $vm_base_path)
Write-Host -ForegroundColor $verbose_color "> The new machine will be saved in $vm_base_path."

if (!$accept_defaults) {
    while (!($generation -eq 1) -or (!$generation -eq 2)) {
        $generation = Read-Host "`nChoose generation 1 or 2 (Default: 2)"
        if (!$generation) {
            break
        }
    }
}
if (!$generation) {
    $generation = "2"
}
$parameters.Add("Generation", $generation)
Write-Host -ForegroundColor $verbose_color "> The new machine will be of generation $generation."

if (!$accept_defaults) {
    $memory = Read-Host "`nDefine RAM size (Default: 2GB)"
}
if (!$memory) {
    $memory = "2GB"
}
$parameters.Add("MemoryStartupBytes", $memory)
Write-Host -ForegroundColor $verbose_color "> The new machine will have $memory of memory."

$vm_switch = Get-VMSwitch -SwitchType External
$vm_switch_name = $vm_switch.Name
if (!$accept_defaults) {
    $switch_choice = Read-Host "`nDo you want to use $vm_switch_name with this machine? ([Y]es/[N]o/[P]ick Default: Y)"
}
if ($switch_choice -like "N*") {
    Write-Host -ForegroundColor $verbose_color "> No network switch will be attached to this machine."
} elseif ($switch_choice -like "P*") {
    $vm_switch_name = PickSwitch
    $parameters.Add("SwitchName", $vm_switch_name)
    Write-Host -ForegroundColor $verbose_color "> $vm_switch_name will be used with this machine."
} else {
    $parameters.Add("SwitchName", $vm_switch_name)
    Write-Host -ForegroundColor $verbose_color "> $vm_switch_name will be used with this machine."
}

if (!$accept_defaults) {
    $disk_choice = Read-Host "`nDo you want to create a new disk for this machine? ([Y]es/[N]o Default: Y)"
}
if ($disk_choice -like "N*") {
    $parameters.Add("NoVHD", $true)
    Write-Host -ForegroundColor $verbose_color "> The new machine will be created without a virtual disk."
} else {
    $disk_path = "$vm_base_path\$vm_name\Virtual Hard Disks\$vm_name.vhdx"
    if (!$accept_defaults) {
        $vm_disk_size = Read-Host "Choose a size for your disk (Default: 40GB)"
    }
    if (!$vm_disk_size) {
        $vm_disk_size = "40GB"
    }
    $parameters.Add("NewVHDPath", $disk_path)
    $parameters.Add("NewVHDSizeBytes", $vm_disk_size)
    Write-Host -ForegroundColor $verbose_color "> A new disk of $vm_disk_size will be created in $disk_path."
}

# Create VM
New-VM @parameters > $null
if ($?) {
    Write-Host -ForegroundColor $success_color "> $vm_name was successfully created."
}

# Configure VM
if (!$accept_defaults) {
    $secure_boot = Read-Host "`nDo you want to disable secure boot? ([Y]es/[N]o Default: N)"
}
if ($secure_boot -like "Y*") {
    Set-VMFirmware -VMName $vm_name -EnableSecureBoot Off
    Write-Host -ForegroundColor $verbose_color "> Secure boot was disabled on $vm_name."
}

if (!$accept_defaults) {
    $cores = Read-Host "`nDo you want to increase core count? (Default: 2)"
}
if (!$cores) {
    $cores = "2"
}
Set-VM -Name $vm_name -ProcessorCount $cores
Write-Host -ForegroundColor $verbose_color "> $vm_name was assigned $cores cores."

if (!$accept_defaults) {
    $snapshots = Read-Host "`nDo you want to disable snapshots? ([Y]es/[N]o Default: Y)"
}
if (!$snapshots -or $snapshots -like "Y*") {
    Set-VM -Name $vm_name -CheckpointType Disabled
    Write-Host -ForegroundColor $verbose_color "> Snapshots were disabled on $vm_name.`n"
}

if (!$accept_defaults) {
    $iso_choice = Read-Host "Do you want to add an ISO to this machine? ([Y]es/[N]o Default: N)"
    if ($iso_choice -like "Y*") {
        $iso_path = Read-Host "Press [ENTER] to choose a file or type a path"
        if (!$iso_path) {
            $iso_path = PickFolder -Object System.Windows.Forms.OpenFileDialog
        }
        Add-VMDvdDrive -VMName $vm_name -ControllerNumber 0 -ControllerLocation 1 -Path $iso_path
        $scsi_dvd = Get-VMDvdDrive -VMName $vm_name
        Set-VMFirmware -VMName $vm_name -FirstBootDevice $scsi_dvd
        Write-Host -ForegroundColor $success_color "> $iso_path was successfully added to the machine.`n"
    } else {
        Write-Host "> Goodbye.`n"
    }
}